#!/usr/bin/env python3
"""
extract_attachments.py

Extrait les pieces jointes des emails (envoi comme reception, dans toute la
boite) de plusieurs comptes Microsoft 365 donnes, via Microsoft Graph
(app-only / client credentials), et les depose (optionnellement) dans un
dossier SharePoint donne par un lien de partage.

Chaque adresse listee est interrogee comme sa PROPRE boite mail (pas de
boite centrale) : /users/{adresse}/messages. Le filtre par mot-cle (ex:
"OV" pour Ordre de Virement) s'applique ensuite sur l'objet de l'email et
le nom des pieces jointes.

Reutilise l'App Registration Entra ID "BackupOffice365" (deja utilisee pour
l'automatisation mWater -> SharePoint), a laquelle il faut avoir ajoute la
permission Application "Mail.Read" (ou "Mail.ReadBasic.All"), avec
consentement admin accorde dans le portail Entra ID. La permission
Application "Sites.ReadWrite.All" (deja presente pour le backup mWater)
est reutilisee pour l'upload SharePoint.

Authentification via les memes secrets que le script de backup mWater :
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET

Le token app-only expire au bout d'environ 1h. Comme un run peut durer plus
longtemps (beaucoup de boites/pieces jointes), le token est rafraichi
automatiquement avant expiration et sur toute reponse 401 en cours de route.

Optionnel, pour deposer les fichiers sur SharePoint en plus du disque local :
    SHAREPOINT_FOLDER_LINK   (lien de partage du dossier cible, type
                               https://xxx.sharepoint.com/:f:/s/.../...)

Usage :
    python extract_attachments.py \
        --senders mickael.consultant@madavance.org rakitrynyavo@madavance.org holisoa.raharijaona@madavance.org

    # Une seule boite, filtre desactive :
    python extract_attachments.py \
        --senders rakitrynyavo@madavance.org \
        --keyword "" \
        --output-dir ./pieces_jointes
"""

import argparse
import base64
import hashlib
import os
import re
import sys
import time
import unicodedata
from pathlib import Path
from urllib.parse import quote

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SENDERS = [
    "mickael.consultant@madavance.org",
    "rakitrynyavo@madavance.org",
    "holisoa.raharijaona@madavance.org",
]
DEFAULT_KEYWORD = "OV"
# Taille de chunk pour l'upload SharePoint : doit etre un multiple de 320 KiB.
CHUNK_SIZE = 320 * 1024 * 30  # ~9,37 Mo
# Marge de securite avant expiration du token pour declencher un refresh proactif.
TOKEN_REFRESH_MARGIN_SECONDS = 120


def raise_for_status_verbose(resp: requests.Response) -> None:
    """Leve une exception avec le corps de la reponse en cas d'erreur HTTP."""
    if resp.status_code >= 400:
        raise RuntimeError(
            f"Erreur HTTP {resp.status_code} sur {resp.request.method} {resp.url}\n{resp.text}"
        )


class GraphSession:
    """Gere le token app-only (client credentials) et le rafraichit automatiquement :
    - de maniere proactive, avant qu'il n'expire (marge de securite) ;
    - de maniere reactive, si un appel renvoie quand meme 401 (horloge, latence...).
    Toutes les requetes Graph du script passent par ici plutot que par un token
    brut, pour eviter le crash "token is expired" en plein milieu d'un run long."""

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
        """Force une premiere authentification (echoue vite si les secrets sont faux)."""
        self._token_value()

    def request(self, method: str, url: str, **kwargs) -> requests.Response:
        headers = dict(kwargs.pop("headers", None) or {})
        headers["Authorization"] = f"Bearer {self._token_value()}"
        resp = requests.request(method, url, headers=headers, **kwargs)
        if resp.status_code == 401:
            # Le token a pu expirer entre le check proactif et l'appel reel
            # (run long, latence reseau...). On force un refresh et on retente une fois.
            headers["Authorization"] = f"Bearer {self._token_value(force_refresh=True)}"
            resp = requests.request(method, url, headers=headers, **kwargs)
        return resp

    def get(self, url: str, **kwargs) -> requests.Response:
        return self.request("GET", url, **kwargs)

    def post(self, url: str, **kwargs) -> requests.Response:
        return self.request("POST", url, **kwargs)

    def put(self, url: str, **kwargs) -> requests.Response:
        return self.request("PUT", url, **kwargs)


def list_all_messages_with_attachments(session: GraphSession, mailbox: str) -> list[dict]:
    """Liste tous les messages avec pieces jointes de la boite donnee
    (envoi + reception confondus, car /users/{id}/messages couvre toute
    la boite, pas seulement la reception)."""
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {
        "$filter": "hasAttachments eq true",
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
        "$top": "100",
    }
    # Note : ne pas combiner $filter et $orderby ici -> Graph renvoie
    # "InefficientFilter" (400) sur /messages avec ce type de filtre.
    messages = []
    while url:
        resp = session.get(url, params=params, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        params = None  # nextLink embarque deja les query params
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
    """Nettoie une chaine pour en faire un nom de dossier/fichier sur."""
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^\w\-. ]", "_", value).strip()
    value = re.sub(r"\s+", "_", value)
    return value[:max_len] or "sans_nom"


def make_keyword_matcher(keyword: str):
    """Construit un matcher qui reconnait un mot-cle en tant que token entier
    (ex: 'OV', 'OV123', 'ov_2024'), pour eviter les faux positifs du type
    'novembre' contenant la sous-chaine 'ov'. Si keyword est vide, tout matche."""
    if not keyword:
        return lambda text: True
    pattern = re.compile(rf"^{re.escape(keyword)}\d*$", re.IGNORECASE)

    def matcher(text: str) -> bool:
        if not text:
            return False
        tokens = re.split(r"[^A-Za-z0-9]+", text)
        return any(pattern.match(t) for t in tokens if t)

    return matcher


def build_filename(received: str, subject: str, original_name: str, max_stem_len: int = 60) -> str:
    """Construit un nom de fichier sur, en preservant toujours la vraie extension
    d'origine (contrairement a une simple troncature qui peut, par coincidence,
    couper juste apres un '.' et faire croire qu'une extension est deja presente
    alors qu'elle a ete tronquee -> fichier sans extension, rejete par SharePoint)."""
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
    """Evite d'ecraser un fichier existant en ajoutant un suffixe numerique."""
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
    """Encode une URL de partage au format attendu par /shares/{id}."""
    b64 = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{b64}"


def resolve_share_link(session: GraphSession, share_url: str) -> tuple[str, str]:
    """Resout un lien de partage SharePoint en (driveId, itemId) du dossier cible."""
    encoded = encode_sharing_url(share_url)
    url = f"{GRAPH_BASE}/shares/{encoded}/driveItem"
    resp = session.get(url, timeout=30)
    raise_for_status_verbose(resp)
    item = resp.json()
    drive_id = item["parentReference"]["driveId"]
    item_id = item["id"]
    return drive_id, item_id


def get_or_create_child_folder(session: GraphSession, drive_id: str, parent_item_id: str, name: str) -> str:
    """Trouve un sous-dossier par nom sous un item donne, ou le cree s'il n'existe pas."""
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
    """Retourne l'id du sous-dossier {annee}/{mois} sous base_folder_id, en le
    creant si besoin. Les resultats sont mis en cache pour eviter de refaire
    les memes appels Graph pour chaque piece jointe du meme mois."""
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
    """Upload un fichier (petit ou volumineux) dans un dossier SharePoint via upload session."""
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
        # L'upload session Graph a sa propre URL pre-signee : pas besoin (ni
        # souhaitable) d'y rajouter le header Authorization app-only.
        put_headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {start}-{end}/{total}",
        }
        resp = requests.put(upload_url, headers=put_headers, data=chunk, timeout=120)
        raise_for_status_verbose(resp)
        start = end + 1


# --- Programme principal ----------------------------------------------------------


def main() -> int:
    parser = argparse.ArgumentParser(description="Extrait les pieces jointes de plusieurs boites mail donnees.")
    parser.add_argument(
        "--senders",
        nargs="+",
        default=DEFAULT_SENDERS,
        help=f"Adresses email dont on interroge directement la boite (defaut: {', '.join(DEFAULT_SENDERS)})",
    )
    parser.add_argument("--output-dir", default="./pieces_jointes", help="Dossier de sortie local")
    parser.add_argument(
        "--keyword",
        default=os.environ.get("FILTER_KEYWORD", DEFAULT_KEYWORD),
        help=(
            f"Mot-cle a rechercher dans l'objet de l'email OU le nom du fichier joint "
            f"(defaut: {DEFAULT_KEYWORD}). Chaine vide pour desactiver le filtre."
        ),
    )
    parser.add_argument(
        "--sharepoint-link",
        default=os.environ.get("SHAREPOINT_FOLDER_LINK"),
        help="Lien de partage du dossier SharePoint cible (optionnel, sinon variable SHAREPOINT_FOLDER_LINK)",
    )
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID et AZURE_CLIENT_SECRET doivent etre definis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    session = GraphSession(tenant_id, client_id, client_secret)
    session.ensure_ready()

    keyword_matches = make_keyword_matcher(args.keyword)
    if args.keyword:
        print(f"Filtre actif : objet OU nom de fichier contenant '{args.keyword}'.")

    sp_drive_id = sp_folder_id = None
    if args.sharepoint_link:
        print("Resolution du dossier SharePoint cible...")
        sp_drive_id, sp_folder_id = resolve_share_link(session, args.sharepoint_link)
        print(f"  -> driveId={sp_drive_id} folderId={sp_folder_id}")

    output_root = Path(args.output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    # Classement par annee/mois (date de reception) : plus de sous-dossier par
    # boite, mais un sous-dossier {annee}/{mois} sous le dossier cible. Le
    # dedoublonnage (par contenu, via sha256) evite d'enregistrer/uploader deux
    # fois le meme fichier quand il a ete echange entre plusieurs boites.
    seen_hashes: set = set()
    sp_folder_cache: dict[tuple[str, str], str] = {}
    total_files = 0
    total_duplicates = 0
    for mailbox in args.senders:
        print(f"\n--- Boite : {mailbox} ---")
        messages = list_all_messages_with_attachments(session, mailbox)
        print(f"{len(messages)} email(s) avec piece(s) jointe(s) dans cette boite (envoi + reception).")

        for msg in messages:
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]
            attachments = list_attachments(session, mailbox, msg["id"])
            file_attachments = [a for a in attachments if a.get("@odata.type") == "#microsoft.graph.fileAttachment"]

            if not file_attachments:
                continue

            # Filet de securite : en plus du token exact ("OV", "OV123"...), on
            # declenche aussi sur "virement" dans l'objet (ex: emails intitules
            # "2 Ordres de virement" sans le sigle "OV"). Les pieces jointes
            # elles-memes suivent toujours la convention "OV <Nom>_<date>.ext".
            subject_matches = keyword_matches(subject) or "virement" in subject.lower()
            kept_attachments = [
                att for att in file_attachments
                if subject_matches or keyword_matches(att.get("name") or "")
            ]
            if not kept_attachments:
                continue

            print(f"  [{received}] {subject} - {len(kept_attachments)}/{len(file_attachments)} piece(s) jointe(s) retenue(s)")

            # Classement par annee/mois de reception (ex: 2025/11). Repli sur
            # "date_inconnue" si jamais receivedDateTime est absent.
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
                print(f"    -> local: {dest}")

                if sp_folder_id:
                    sp_month_folder_id = get_year_month_folder(session, sp_drive_id, sp_folder_id, year, month, sp_folder_cache)
                    upload_file_to_sharepoint(session, sp_drive_id, sp_month_folder_id, filename, content)
                    print(f"    -> SharePoint: {year}/{month}/{filename}")

    print(f"\nTermine. {total_files} piece(s) jointe(s) enregistree(s) dans {output_root.resolve()}")
    if total_duplicates:
        print(f"{total_duplicates} doublon(s) detecte(s) et ignore(s) (meme contenu deja enregistre).")
    if sp_drive_id:
        print("Egalement deposees sur SharePoint, classees par sous-dossiers annee/mois (pas de sous-dossier par boite).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
