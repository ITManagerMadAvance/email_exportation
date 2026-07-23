#!/usr/bin/env python3
"""
extract_invoices.py -- "Extraction Factures / Proforma"

Meme principe que extract_attachments.py (Extraction OV) : interroge
directement plusieurs boites Microsoft 365 (envoi + reception), et en
extrait les pieces jointes filtrees par mot-cle. Ce script est cependant
un flow SEPARE de l'extraction OV -- ne modifie pas extract_attachments.py,
ne partage pas son dossier de destination.

Deux modes, lances TOUJOURS EN MEME TEMPS dans une seule execution (pas de
choix a faire au lancement) :
    facture   -> mot-cles "facture" / "invoice"
    proforma  -> mot-cles "devis" / "proforma"

Chaque mode a son propre dossier SharePoint de destination (jamais melanges,
meme s'ils sont extraits dans la meme execution) :
    SHAREPOINT_FOLDER_LINK_FACTURES   (mode facture)
    SHAREPOINT_FOLDER_LINK_PROFORMA   (mode proforma)

--mode reste disponible en option pour ne lancer qu'un seul des deux modes
(utile pour tester), mais le comportement par defaut (aucun --mode fourni)
est de faire les deux a la suite, avec un seul email de synthese couvrant
les deux.

Le filtre s'applique sur l'objet de l'email OU le nom de la piece jointe,
en reconnaissant le mot-cle comme token entier (avec suffixe chiffre
optionnel, ex: "facture", "FACTURE2026", "devis_12"), pour eviter les faux
positifs de sous-chaine.

Faux positifs ecartes volontairement (observes en pratique dans
admin@madavance.org) : releves bancaires ("RELEVE BQ..."), documents
administratifs/legaux (certificats, carte fiscale, CIF), etat de paie,
journal de caisse -- aucun de ces mots-cles n'apparait dans "facture" ou
"proforma"/"devis", donc ils sont deja naturellement exclus par le filtre
token-exact.

Authentification identique aux autres scripts du repo (BackupOffice365,
app-only) :
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET

Usage :
    python extract_invoices.py --mode facture
    python extract_invoices.py --mode proforma --senders olivia@madavance.org
"""

import argparse
import base64
import hashlib
import os
import re
import sys
import time
import traceback
import unicodedata
from datetime import date
from pathlib import Path
from urllib.parse import quote

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SENDERS = [
    "mickael.consultant@madavance.org",
    "rakitrynyavo@madavance.org",
    "holisoa.raharijaona@madavance.org",
    "olivia@madavance.org",
]
MODE_KEYWORDS = {
    "facture": ["facture", "invoice"],
    "proforma": ["devis", "proforma"],
}
CHUNK_SIZE = 320 * 1024 * 30  # ~9,37 Mo
TOKEN_REFRESH_MARGIN_SECONDS = 120
TRANSIENT_STATUS_CODES = {429, 500, 502, 503, 504}
MAX_TRANSIENT_RETRIES = 5
MAX_RETRY_DELAY_SECONDS = 60


def raise_for_status_verbose(resp: requests.Response) -> None:
    if resp.status_code >= 400:
        raise RuntimeError(
            f"Erreur HTTP {resp.status_code} sur {resp.request.method} {resp.url}\n{resp.text}"
        )


def _compute_retry_delay(resp: requests.Response, attempt: int) -> float:
    retry_after = resp.headers.get("Retry-After")
    if retry_after:
        try:
            return float(retry_after)
        except ValueError:
            pass
    try:
        retry_seconds = resp.json().get("error", {}).get("retryAfterSeconds")
        if retry_seconds:
            return float(retry_seconds)
    except Exception:
        pass
    return min(2 ** attempt, MAX_RETRY_DELAY_SECONDS)


def request_with_retry(method: str, url: str, **kwargs) -> requests.Response:
    attempt = 0
    while True:
        resp = requests.request(method, url, **kwargs)
        if resp.status_code not in TRANSIENT_STATUS_CODES or attempt >= MAX_TRANSIENT_RETRIES:
            return resp
        delay = _compute_retry_delay(resp, attempt)
        print(f"    (Graph a renvoye {resp.status_code}, nouvelle tentative dans {delay:.0f}s...)")
        time.sleep(delay)
        attempt += 1


class GraphSession:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self._token = None
        self._expires_at = 0

    def _fetch_token(self) -> None:
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        resp = requests.post(url, data=data, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        self._token = payload["access_token"]
        expires_in = int(payload.get("expires_in", 3599))
        self._expires_at = time.time() + expires_in

    def _token_value(self, force_refresh: bool = False) -> str:
        if force_refresh or self._token is None or time.time() >= self._expires_at - TOKEN_REFRESH_MARGIN_SECONDS:
            self._fetch_token()
        return self._token

    def ensure_ready(self) -> None:
        self._token_value()

    def request(self, method: str, url: str, **kwargs) -> requests.Response:
        headers = dict(kwargs.pop("headers", None) or {})
        headers["Authorization"] = f"Bearer {self._token_value()}"
        resp = request_with_retry(method, url, headers=headers, **kwargs)
        if resp.status_code == 401:
            headers["Authorization"] = f"Bearer {self._token_value(force_refresh=True)}"
            resp = request_with_retry(method, url, headers=headers, **kwargs)
        return resp

    def get(self, url: str, **kwargs) -> requests.Response:
        return self.request("GET", url, **kwargs)

    def post(self, url: str, **kwargs) -> requests.Response:
        return self.request("POST", url, **kwargs)

    def put(self, url: str, **kwargs) -> requests.Response:
        return self.request("PUT", url, **kwargs)


def list_all_messages_with_attachments(session: GraphSession, mailbox: str) -> list[dict]:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {
        "$filter": "hasAttachments eq true",
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
        "$top": "100",
    }
    messages = []
    while url:
        resp = session.get(url, params=params, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        params = None
    return messages


def list_attachments(session: GraphSession, mailbox: str, message_id: str) -> list[dict]:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    resp = session.get(url, timeout=30)
    raise_for_status_verbose(resp)
    return resp.json().get("value", [])


def download_attachment_bytes(session: GraphSession, mailbox: str, message_id: str, attachment_id: str) -> bytes:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments/{attachment_id}/$value"
    resp = session.get(url, timeout=120)
    raise_for_status_verbose(resp)
    return resp.content


def slugify(value: str, max_len: int = 60) -> str:
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^\w\-. ]", "_", value).strip()
    value = re.sub(r"\s+", "_", value)
    return value[:max_len] or "sans_nom"


def make_keyword_matcher(keywords: list[str]):
    """Comme dans extract_attachments.py, mais accepte plusieurs mots-cles
    (ex: 'facture' et 'invoice'). Chaque mot-cle est reconnu comme token
    entier avec suffixe chiffre optionnel (ex: 'facture', 'FACTURE2026')."""
    keywords = [k for k in keywords if k]
    if not keywords:
        return lambda text: True
    patterns = [re.compile(rf"^{re.escape(k)}\d*$", re.IGNORECASE) for k in keywords]

    def matcher(text: str) -> bool:
        if not text:
            return False
        tokens = re.split(r"[^A-Za-z0-9]+", text)
        return any(p.match(t) for t in tokens if t for p in patterns)

    return matcher


def build_filename(received: str, subject: str, original_name: str, max_stem_len: int = 60) -> str:
    base_name, dot, ext = original_name.rpartition(".")
    stem_source = base_name if dot else original_name
    stem = slugify(stem_source, max_stem_len)
    prefix = f"{received}_{slugify(subject, 40)}_{stem}"
    if dot and ext:
        ext_clean = re.sub(r"[^A-Za-z0-9]", "", ext)[:10]
        if ext_clean:
            return f"{prefix}.{ext_clean.lower()}"
    return prefix


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    i = 2
    while True:
        candidate = path.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


# --- SharePoint (Microsoft Graph) ------------------------------------------------


def encode_sharing_url(url: str) -> str:
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{b64}"


def resolve_share_link(session: GraphSession, share_url: str) -> tuple[str, str]:
    encoded = encode_sharing_url(share_url)
    url = f"{GRAPH_BASE}/shares/{encoded}/driveItem"
    resp = session.get(url, timeout=30)
    raise_for_status_verbose(resp)
    item = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    return drive_id, item_id


def get_or_create_child_folder(session: GraphSession, drive_id: str, parent_item_id: str, name: str) -> str:
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}/children"
    resp = session.get(url, params={"$select": "id,name,folder"}, timeout=30)
    raise_for_status_verbose(resp)
    for item in resp.json().get("value", []):
        if item.get("folder") is not None and item.get("name") == name:
            return item["id"]

    resp = session.post(
        url,
        headers={"Content-Type": "application/json"},
        json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "rename"},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["id"]


def get_year_month_folder(
    session: GraphSession,
    drive_id: str,
    base_folder_id: str,
    year: str,
    month: str,
    cache: dict,
) -> str:
    key = (year, month)
    if key in cache:
        return cache[key]

    year_key = (year, "")
    year_folder_id = cache.get(year_key)
    if year_folder_id is None:
        year_folder_id = get_or_create_child_folder(session, drive_id, base_folder_id, year)
        cache[year_key] = year_folder_id

    month_folder_id = get_or_create_child_folder(session, drive_id, year_folder_id, month)
    cache[key] = month_folder_id
    return month_folder_id


def upload_file_to_sharepoint(session: GraphSession, drive_id: str, parent_item_id: str, filename: str, content: bytes) -> None:
    safe_name = quote(filename)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_item_id}:/{safe_name}:/createUploadSession"
    resp = session.post(
        url,
        headers={"Content-Type": "application/json"},
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    upload_url = resp.json()["uploadUrl"]

    total = len(content)
    start = 0
    while start < total:
        end = min(start + CHUNK_SIZE, total) - 1
        chunk = content[start:end + 1]
        put_headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {start}-{end}/{total}",
        }
        resp = request_with_retry("PUT", upload_url, headers=put_headers, data=chunk, timeout=120)
        raise_for_status_verbose(resp)
        start = end + 1


# --- Email de confirmation (Microsoft Graph) --------------------------------------


def send_email(session: GraphSession, sender: str, recipients: str, subject: str, body_text: str) -> None:
    to_recipients = [
        {"emailAddress": {"address": addr.strip()}}
        for addr in recipients.split(",")
        if addr.strip()
    ]
    if not to_recipients:
        print("Aucun destinataire valide dans EMAIL_RECIPIENTS, email non envoye.")
        return

    url = f"{GRAPH_BASE}/users/{sender}/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": to_recipients,
        },
        "saveToSentItems": "true",
    }
    resp = session.post(url, json=payload, timeout=30)
    raise_for_status_verbose(resp)
    print(f"Email envoye a {recipients} depuis {sender}.")


# --- Programme principal ----------------------------------------------------------


def run_extraction(session: GraphSession, mode: str, senders: list[str], output_dir: str, sharepoint_link: str | None) -> dict:
    keywords = MODE_KEYWORDS[mode]
    keyword_matches = make_keyword_matcher(keywords)
    print(f"Mode : {mode}. Filtre actif : objet OU nom de fichier contenant {keywords}.")

    sp_drive_id = sp_folder_id = None
    if sharepoint_link:
        print("Resolution du dossier SharePoint cible...")
        sp_drive_id, sp_folder_id = resolve_share_link(session, sharepoint_link)
        print(f"  -> driveId={sp_drive_id} folderId={sp_folder_id}")

    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    seen_hashes: set = set()
    sp_folder_cache: dict[tuple[str, str], str] = {}
    total_files = 0
    total_duplicates = 0
    per_mailbox_counts: dict[str, int] = {}

    for mailbox in senders:
        print(f"\n--- Boite : {mailbox} ---")
        messages = list_all_messages_with_attachments(session, mailbox)
        print(f"{len(messages)} email(s) avec piece(s) jointe(s) dans cette boite (envoi + reception).")
        per_mailbox_counts[mailbox] = 0

        for msg in messages:
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]
            attachments = list_attachments(session, mailbox, msg["id"])
            file_attachments = [a for a in attachments if a.get("@odata.type") == "#microsoft.graph.fileAttachment"]

            if not file_attachments:
                continue

            subject_matches = keyword_matches(subject)
            kept_attachments = [
                att for att in file_attachments
                if subject_matches or keyword_matches(att.get("name") or "")
            ]
            if not kept_attachments:
                continue

            print(f"  [{received}] {subject} - {len(kept_attachments)}/{len(file_attachments)} piece(s) jointe(s) retenue(s)")

            year = received[:4] if len(received) >= 7 else "date_inconnue"
            month = received[5:7] if len(received) >= 7 else "date_inconnue"

            for att in kept_attachments:
                name = att.get("name") or f"piece_jointe_{att['id']}"
                filename = build_filename(received, subject, name)

                content = download_attachment_bytes(session, mailbox, msg["id"], att["id"])
                content_hash = hashlib.sha256(content).hexdigest()

                if content_hash in seen_hashes:
                    total_duplicates += 1
                    print(f"    -> doublon ignore (deja recupere ailleurs): {name}")
                    continue
                seen_hashes.add(content_hash)

                local_dir = output_root / year / month
                local_dir.mkdir(parents=True, exist_ok=True)
                dest = unique_path(local_dir / filename)
                dest.write_bytes(content)
                total_files += 1
                per_mailbox_counts[mailbox] += 1
                print(f"    -> local: {dest}")

                if sp_folder_id:
                    sp_month_folder_id = get_year_month_folder(session, sp_drive_id, sp_folder_id, year, month, sp_folder_cache)
                    upload_file_to_sharepoint(session, sp_drive_id, sp_month_folder_id, filename, content)
                    print(f"    -> SharePoint: {year}/{month}/{filename}")

    print(f"\nTermine. {total_files} piece(s) jointe(s) enregistree(s) dans {output_root.resolve()}")
    if total_duplicates:
        print(f"{total_duplicates} doublon(s) detecte(s) et ignore(s) (meme contenu deja enregistre).")
    if sp_drive_id:
        print("Egalement deposees sur SharePoint, classees par sous-dossiers annee/mois.")

    return {
        "mode": mode,
        "total_files": total_files,
        "total_duplicates": total_duplicates,
        "per_mailbox_counts": per_mailbox_counts,
        "sharepoint_used": bool(sp_drive_id),
    }


def build_mode_summary_text(stats: dict) -> str:
    label = "Factures" if stats["mode"] == "facture" else "Proforma/Devis"
    lines = [f"--- {label} ---"]
    lines.append(f"Total : {stats['total_files']} piece(s) jointe(s) enregistree(s).")
    if stats["total_duplicates"]:
        lines.append(f"Doublons ignores : {stats['total_duplicates']}.")
    lines.append("Detail par boite :")
    for mailbox, count in stats["per_mailbox_counts"].items():
        lines.append(f"  - {mailbox} : {count} piece(s) jointe(s)")
    if stats["sharepoint_used"]:
        lines.append("Deposees sur SharePoint (classees par annee/mois).")
    return "\n".join(lines)


def build_summary_text(all_stats: list[dict]) -> str:
    lines = [f"Extraction Factures / Proforma terminee le {date.today().isoformat()}.", ""]
    for stats in all_stats:
        lines.append(build_mode_summary_text(stats))
        lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extraction Factures + Proforma (les deux en une seule execution, meme lancement, dossiers SharePoint separes)."
    )
    parser.add_argument(
        "--mode",
        choices=["facture", "proforma"],
        default=None,
        help="Ne lancer qu'un seul mode (utile pour tester). Par defaut : les deux modes sont lances a la suite.",
    )
    parser.add_argument(
        "--senders",
        nargs="+",
        default=DEFAULT_SENDERS,
        help=f"Adresses email dont on interroge directement la boite (defaut: {', '.join(DEFAULT_SENDERS)})",
    )
    parser.add_argument("--email-sender", default=os.environ.get("EMAIL_SENDER"))
    parser.add_argument("--email-recipients", default=os.environ.get("EMAIL_RECIPIENTS"))
    args = parser.parse_args()

    modes_to_run = [args.mode] if args.mode else ["facture", "proforma"]

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID et AZURE_CLIENT_SECRET doivent etre definis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    session = GraphSession(tenant_id, client_id, client_secret)
    session.ensure_ready()

    email_enabled = bool(args.email_sender and args.email_recipients)
    if not email_enabled:
        print("EMAIL_SENDER / EMAIL_RECIPIENTS non definis : pas d'email de confirmation envoye.")

    all_stats: list[dict] = []
    errors: list[str] = []

    for mode in modes_to_run:
        print(f"\n=== Mode : {mode} ===")
        output_dir = f"./{mode}"
        env_var = "SHAREPOINT_FOLDER_LINK_FACTURES" if mode == "facture" else "SHAREPOINT_FOLDER_LINK_PROFORMA"
        sharepoint_link = os.environ.get(env_var)
        try:
            stats = run_extraction(session, mode, args.senders, output_dir, sharepoint_link)
            all_stats.append(stats)
        except Exception as exc:
            error_text = f"{exc}\n\n{traceback.format_exc()}"
            print(f"ERREUR (mode {mode}) : {exc}", file=sys.stderr)
            errors.append(f"--- {mode} : ECHEC ---\n{error_text}")

    if email_enabled:
        total_files = sum(s["total_files"] for s in all_stats)
        subject_status = "ECHEC PARTIEL" if errors and all_stats else ("ECHEC" if errors else "OK")
        subject = f"[{subject_status}] Extraction Factures/Proforma - {date.today().isoformat()} : {total_files} fichier(s)"
        body = build_summary_text(all_stats) if all_stats else ""
        if errors:
            body += "\n\n" + "\n\n".join(errors)
        try:
            send_email(session, args.email_sender, args.email_recipients, subject=subject, body_text=body)
        except Exception as mail_exc:
            print(f"Echec de l'envoi de l'email de confirmation : {mail_exc}", file=sys.stderr)

    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
