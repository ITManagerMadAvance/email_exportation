#!/usr/bin/env python3
"""
list_subjects.py -- Diagnostic en lecture seule

Liste, pour une ou plusieurs boites mail donnees, l'objet de chaque email
ayant au moins une piece jointe, ainsi que le nom de chaque piece jointe.
NE TELECHARGE RIEN, ne filtre par aucun mot-cle, ne modifie rien.

But : observer les vraies donnees (objets, noms de fichiers) avant de definir
les criteres de detection d'une "facture", plutot que de deviner un mot-cle
a l'aveugle.

Authentification identique aux autres scripts du repo (BackupOffice365,
app-only) :
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET

Usage :
    python list_subjects.py --senders mickael.consultant@madavance.org rakitrynyavo@madavance.org holisoa.raharijaona@madavance.org
"""

import argparse
import os
import sys

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SENDERS = [
    "mickael.consultant@madavance.org",
    "rakitrynyavo@madavance.org",
    "holisoa.raharijaona@madavance.org",
]


def get_app_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def list_messages_with_attachments(token: str, mailbox: str) -> list[dict]:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages"
    params = {
        "$filter": "hasAttachments eq true",
        "$select": "id,subject,from,receivedDateTime",
        "$top": "100",
    }
    headers = {"Authorization": f"Bearer {token}"}
    messages = []
    while url:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code >= 400:
            print(f"  ERREUR {resp.status_code} sur {url} : {resp.text}", file=sys.stderr)
            break
        payload = resp.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
        params = None
    return messages


def list_attachment_names(token: str, mailbox: str, message_id: str) -> list[str]:
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/attachments"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params={"$select": "name,contentType"}, timeout=30)
    if resp.status_code >= 400:
        print(f"    ERREUR {resp.status_code} attachments: {resp.text}", file=sys.stderr)
        return []
    return [f"{a.get('name')} ({a.get('contentType')})" for a in resp.json().get("value", [])]


def main() -> int:
    parser = argparse.ArgumentParser(description="Liste (sans telecharger) les objets et noms de pieces jointes.")
    parser.add_argument("--senders", nargs="+", default=DEFAULT_SENDERS)
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET requis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    token = get_app_token(tenant_id, client_id, client_secret)

    for mailbox in args.senders:
        print(f"\n=== Boite : {mailbox} ===")
        messages = list_messages_with_attachments(token, mailbox)
        print(f"{len(messages)} email(s) avec piece(s) jointe(s).")
        for msg in messages:
            subject = msg.get("subject") or "(sans objet)"
            received = (msg.get("receivedDateTime") or "")[:10]
            names = list_attachment_names(token, mailbox, msg["id"])
            print(f"  [{received}] OBJET: {subject}")
            for n in names:
                print(f"        PJ: {n}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
