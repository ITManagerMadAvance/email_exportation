#!/usr/bin/env python3
"""
extract_attachments.py

Extrait les pieces jointes des emails recus d'expediteurs donnes, dans une
boite Microsoft 365, via Microsoft Graph (app-only / client credentials).

Reutilise l'App Registration Entra ID "BackupOffice365" (deja utilisee pour
l'automatisation mWater -> SharePoint), a laquelle il faut avoir ajoute la
permission Application "Mail.Read" (ou "Mail.ReadBasic.All"), avec
consentement admin accorde dans le portail Entra ID.

Authentification via les memes secrets que le script de backup mWater :
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET

Usage :
    python extract_attachments.py \
        --senders mickael.consultant@madavance.org rakitrynyavo@madavance.org holisoa.raharijaona@madavance.org

    # Boite differente ou dossier de sortie different :
    python extract_attachments.py \
        --mailbox it@madavance.org \
        --senders mickael.consultant@madavance.org \
        --output-dir ./pieces_jointes
"""

import argparse
import os
import re
import sys
import unicodedata
from pathlib import Path

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_MAILBOX = "it@madavance.org"


def raise_for_status_verbose(resp: requests.Response) -> None:
    """Leve une exception avec le corps de la reponse en cas d'erreur HTTP."""
    if resp.status_code >= 400:
        raise RuntimeError(
            f"Erreur HTTP {resp.status_code} sur {resp.request.method} {resp.url}\n{resp.text}"
        )


def get_app_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Authentification app-only (client credentials) aupres de Microsoft Graph."""
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    raise_for_status_verbose(resp)
    return resp.json()["access_token"]


def list_messages_from_sender(token: str, mailbox: str, sender: str) -> list[dict]:
    """Liste les messages d'un expediteur donne, avec pieces jointes, dans une boite."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {
        "$filter": f"from/emailAddress/address eq '{sender}' and hasAttachments eq true",
        "$select": "id,subject,from,receivedDateTime,hasAttachments",
        "$top": "100",
        "$orderby": "receivedDateTime desc",
    }
    messages = []
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        raise_for_status_verbose(resp)
        payload = resp.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        params = None  # nextLink embarque deja les query params
    return messages


def list_attachments(token: str, mailbox: str, message_id: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    resp = requests.get(url, headers=headers, timeout=30)
    raise_for_status_verbose(resp)
    return resp.json().get("value", [])


def download_attachment_bytes(token: str, mailbox: str, message_id: str, attachment_id: str) -> bytes:
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments/{attachment_id}/$value"
    resp = requests.get(url, headers=headers, timeout=120)
    raise_for_status_verbose(resp)
    return resp.content


def slugify(value: str, max_len: int = 60) -> str:
    """Nettoie une chaine pour en faire un nom de dossier/fichier sur."""
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = re.sub(r"[^\w\-. ]", "_", value).strip()
    value = re.sub(r"\s+", "_", value)
    return value[:max_len] or "sans_nom"


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


def main() -> int:
    parser = argparse.ArgumentParser(description="Extrait les pieces jointes d'expediteurs donnes.")
    parser.add_argument("--mailbox", default=DEFAULT_MAILBOX, help=f"Boite mail a interroger (defaut: {DEFAULT_MAILBOX})")
    parser.add_argument("--senders", nargs="+", required=True, help="Adresses email des expediteurs")
    parser.add_argument("--output-dir", default="./pieces_jointes", help="Dossier de sortie")
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID et AZURE_CLIENT_SECRET doivent etre definis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    token = get_app_token(tenant_id, client_id, client_secret)

    output_root = Path(args.output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    total_files = 0
    for sender in args.senders:
        print(f"\n--- Expediteur : {sender} ---")
        messages = list_messages_from_sender(token, args.mailbox, sender)
        print(f"{len(messages)} email(s) avec piece(s) jointe(s) trouve(s).")

        sender_dir = output_root / slugify(sender)
        sender_dir.mkdir(parents=True, exist_ok=True)

        for msg in messages:
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]
            attachments = list_attachments(token, args.mailbox, msg["id"])
            file_attachments = [a for a in attachments if a.get("@odata.type") == "#microsoft.graph.fileAttachment"]

            if not file_attachments:
                continue

            print(f"  [{received}] {subject} - {len(file_attachments)} piece(s) jointe(s)")

            for att in file_attachments:
                name = att.get("name") or f"piece_jointe_{att['id']}"
                filename = f"{received}_{slugify(subject, 40)}_{slugify(name, 60)}"
                # preserve l'extension d'origine si slugify l'a mangee
                if "." in name and "." not in filename[-6:]:
                    ext = name.rsplit(".", 1)[-1]
                    filename = f"{filename}.{ext}"

                dest = unique_path(sender_dir / filename)
                content = download_attachment_bytes(token, args.mailbox, msg["id"], att["id"])
                dest.write_bytes(content)
                total_files += 1
                print(f"    -> {dest}")

    print(f"\nTermine. {total_files} piece(s) jointe(s) enregistree(s) dans {output_root.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
