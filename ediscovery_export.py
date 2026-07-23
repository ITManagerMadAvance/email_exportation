#!/usr/bin/env python3
"""
ediscovery_export.py

Script DIAGNOSTIC pour verifier l'acces app-only a l'API Microsoft Graph
eDiscovery (/security/cases/ediscoveryCases), et localiser la recherche
"Eddy" deja creee dans le cas Purview "Content Search".

Etape 1 (toujours executee) : liste les cas eDiscovery, liste les recherches
du cas correspondant, affiche les statistiques de la recherche ciblee. Ceci
permet de confirmer que la permission Application "eDiscovery.ReadWrite.All"
(ajoutee sur BackupOffice365) donne bien acces, independamment de la
restriction "Compte Gratuit" vue dans le portail Purview.

Etape 2 (optionnelle, --export) : tente de declencher un export de la
recherche via l'API et d'en suivre la progression. Cette partie est plus
incertaine (l'API d'export eDiscovery a plusieurs variantes selon le type de
cas) -- elle sert a decouvrir ce qui fonctionne reellement sur ce tenant.

Authentification (memes secrets que les autres scripts du repo) :
    AZURE_TENANT_ID
    AZURE_CLIENT_ID
    AZURE_CLIENT_SECRET

Usage :
    python ediscovery_export.py --case-name "Content Search" --search-name "Eddy"
    python ediscovery_export.py --case-name "Content Search" --search-name "Eddy" --export
"""

import argparse
import json
import os
import sys
import time

import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def raise_for_status_verbose(resp: requests.Response) -> None:
    if resp.status_code >= 400:
        raise RuntimeError(
            f"Erreur HTTP {resp.status_code} sur {resp.request.method} {resp.url}\n{resp.text}"
        )


def get_app_token(tenant_id: str, client_id: str, client_secret: str) -> str:
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


def graph_get(token: str, url: str, **kwargs) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, timeout=60, **kwargs)
    raise_for_status_verbose(resp)
    return resp.json()


def graph_post(token: str, url: str, body: dict | None = None) -> requests.Response:
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(url, headers=headers, json=body or {}, timeout=60)
    return resp


def find_case(token: str, case_name: str) -> dict:
    print(f"Recherche du cas eDiscovery '{case_name}'...")
    payload = graph_get(token, f"{GRAPH_BASE}/security/cases/ediscoveryCases")
    cases = payload.get("value", [])
    print(f"{len(cases)} cas trouve(s) au total :")
    for c in cases:
        print(f"  - {c.get('displayName')} (id={c.get('id')}, status={c.get('status')})")
    matches = [c for c in cases if (c.get("displayName") or "").lower() == case_name.lower()]
    if not matches:
        raise RuntimeError(f"Aucun cas nomme '{case_name}' trouve.")
    return matches[0]


def find_search(token: str, case_id: str, search_name: str) -> dict:
    print(f"\nRecherche de la recherche '{search_name}' dans le cas...")
    payload = graph_get(token, f"{GRAPH_BASE}/security/cases/ediscoveryCases/{case_id}/searches")
    searches = payload.get("value", [])
    print(f"{len(searches)} recherche(s) trouvee(s) dans ce cas :")
    for s in searches:
        print(f"  - {s.get('displayName')} (id={s.get('id')})")
    matches = [s for s in searches if (s.get("displayName") or "").lower() == search_name.lower()]
    if not matches:
        raise RuntimeError(f"Aucune recherche nommee '{search_name}' trouvee dans ce cas.")
    return matches[0]


def print_search_details(token: str, case_id: str, search_id: str) -> None:
    print("\nDetails de la recherche :")
    details = graph_get(token, f"{GRAPH_BASE}/security/cases/ediscoveryCases/{case_id}/searches/{search_id}")
    print(json.dumps(details, indent=2, ensure_ascii=False)[:3000])

    print("\nDerniere estimation statistique (si disponible) :")
    try:
        stats = graph_get(
            token,
            f"{GRAPH_BASE}/security/cases/ediscoveryCases/{case_id}/searches/{search_id}/lastEstimateStatisticsOperation",
        )
        print(json.dumps(stats, indent=2, ensure_ascii=False)[:3000])
    except RuntimeError as exc:
        print(f"  (non disponible : {exc})")


def try_export(token: str, case_id: str, search_id: str) -> None:
    print("\n--- Tentative d'export via l'API ---")
    url = f"{GRAPH_BASE}/security/cases/ediscoveryCases/{case_id}/searches/{search_id}/microsoft.graph.security.export"
    body = {
        "description": "Export via script (recuperation contenu Eddy)",
        "outputName": "Export_Eddy",
        "exportSingleItems": True,
        "exportCriteria": "searchHits",
        "exportFormat": "pst",
    }
    resp = graph_post(token, url, body)
    print(f"Statut : {resp.status_code}")
    print(resp.text[:3000])

    if resp.status_code not in (200, 201, 202):
        print("\nL'export via l'API a echoue -- probablement la meme restriction que dans le portail (licence).")
        return

    op_location = resp.headers.get("Location")
    if not op_location:
        print("Pas d'URL d'operation retournee, impossible de suivre la progression automatiquement.")
        return

    print(f"Operation lancee, suivi sur : {op_location}")
    for attempt in range(20):
        time.sleep(15)
        op = graph_get(token, op_location)
        status = op.get("status")
        print(f"  [{attempt+1}] statut : {status}")
        if status in ("succeeded", "failed", "partiallySucceeded"):
            print(json.dumps(op, indent=2, ensure_ascii=False)[:3000])
            break


def main() -> int:
    parser = argparse.ArgumentParser(description="Diagnostic API eDiscovery app-only.")
    parser.add_argument("--case-name", default="Content Search")
    parser.add_argument("--search-name", default="Eddy")
    parser.add_argument("--export", action="store_true", help="Tente en plus de declencher un export via l'API")
    args = parser.parse_args()

    tenant_id = os.environ.get("AZURE_TENANT_ID")
    client_id = os.environ.get("AZURE_CLIENT_ID")
    client_secret = os.environ.get("AZURE_CLIENT_SECRET")
    if not all([tenant_id, client_id, client_secret]):
        print("Erreur : AZURE_TENANT_ID, AZURE_CLIENT_ID et AZURE_CLIENT_SECRET doivent etre definis.", file=sys.stderr)
        return 1

    print("Authentification Microsoft Graph (app-only)...")
    token = get_app_token(tenant_id, client_id, client_secret)

    try:
        case = find_case(token, args.case_name)
    except RuntimeError as exc:
        print(f"ERREUR : {exc}", file=sys.stderr)
        return 1

    case_id = case["id"]

    try:
        search = find_search(token, case_id, args.search_name)
    except RuntimeError as exc:
        print(f"ERREUR : {exc}", file=sys.stderr)
        return 1

    search_id = search["id"]
    print(f"\n=> Cas id={case_id} / Recherche id={search_id}")

    print_search_details(token, case_id, search_id)

    if args.export:
        try_export(token, case_id, search_id)
    else:
        print("\n(Ajoutez --export pour tenter de declencher un export via l'API.)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
