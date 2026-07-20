"""
Rebuild propre du journal ledger depuis les onglets section 1 du sheet.
Chaque ligne = un appel à /api/ledger/sheet-entry → contrepartie PCG réelle.
Pas de réconciliation floue, pas de 599999.
"""

import openpyxl
import requests
import re
from datetime import datetime, date

XLSX = '/home/ubuntu/analyzor/test_copro.xlsx'
API  = 'http://localhost:8080/api/ledger/sheet-entry'
ORG  = 'copro_1crE1G2RerFeXQfHNh0yERfvfAjVKGUz53LE9szCqMMs'

# D'abord on vide le journal (backup automatique dans ledger_api)
RESET_API = 'http://localhost:8080/api/ledger/import'


def parse_amount(v):
    if v is None or v == '':
        return 0.0
    try:
        return float(str(v).replace(',', '.').replace(' ', ''))
    except Exception:
        return 0.0


def parse_date(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y/%m/%d')
    s = str(v).strip()
    if not s or s.lower() in ('date', 'solde', ''):
        return None
    # formats: 01/01/22, 1/1/22, 2022/01/01, 2022-01-01
    for fmt in ('%d/%m/%y', '%d/%m/%Y', '%Y/%m/%d', '%Y-%m-%d', '%m/%d/%y', '%m/%d/%Y'):
        try:
            return datetime.strptime(s, fmt).strftime('%Y/%m/%d')
        except ValueError:
            continue
    return None


def extract_compte(sheet_name):
    """Extrait le code compte depuis le nom de l'onglet '1 - 451001 AMSELLEM' → '451001'"""
    m = re.search(r'\b(\d{3,6})\b', sheet_name)
    return m.group(1) if m else None


wb = openpyxl.load_workbook(XLSX, data_only=True)

ok = 0
skipped = 0
errors = []

for ws in wb.worksheets:
    name = ws.title
    if not name.startswith('1 - '):
        continue
    if name in ('1 - Balance', '1 - SYNTHESE'):
        continue
    if '6 DEPENSES' in name:
        # Structure différente — traité séparément
        continue

    compte = extract_compte(name)
    if not compte:
        print(f'  [SKIP] {name} — compte non trouvé')
        skipped += 1
        continue

    rows = list(ws.iter_rows(values_only=True))

    # Ligne 0 = en-tête compte [code, nom, ...]
    # Ligne 1 = colonnes [Date, Libellé, Débit, Crédit, Solde, ...]
    # Ligne 2+ = données
    data_start = 2

    print(f'\n=== {name} (compte {compte}) ===')
    for i, row in enumerate(rows[data_start:], start=data_start):
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue

        date_val = parse_date(row[0])
        libelle  = str(row[1] or '').strip() if len(row) > 1 else ''
        debit    = parse_amount(row[2]) if len(row) > 2 else 0.0
        credit   = parse_amount(row[3]) if len(row) > 3 else 0.0

        # Ignorer lignes sans date ou libellé (lignes "Solde", totaux, etc.)
        if not date_val or not libelle:
            continue
        if libelle.lower() in ('solde', 'date', 'libellé', 'libelle', ''):
            continue
        if debit == 0 and credit == 0:
            continue

        payload = {
            'orgId': ORG,
            'compte': compte,
            'libelle': libelle,
            'montant_debit': round(debit, 2),
            'montant_credit': round(credit, 2),
            'date': date_val,
        }

        try:
            r = requests.post(API, json=payload, timeout=10)
            res = r.json()
            if res.get('success'):
                cpt = res.get('contrepartie', {})
                print(f'  ✅ {date_val} | {libelle[:35]:<35} | {"D" if debit else "C"} {debit or credit:.2f} → {cpt.get("compte")}')
                ok += 1
            else:
                print(f'  ❌ {date_val} | {libelle[:35]:<35} | {res.get("error")}')
                errors.append({'onglet': name, 'row': i, 'error': res.get('error'), 'payload': payload})
        except Exception as e:
            print(f'  ❌ exception: {e}')
            errors.append({'onglet': name, 'row': i, 'error': str(e), 'payload': payload})

print(f'\n{"="*60}')
print(f'RÉSULTAT : {ok} écritures OK | {skipped} onglets ignorés | {len(errors)} erreurs')
if errors:
    print('\nERREURS :')
    for e in errors[:10]:
        print(f'  {e["onglet"]} ligne {e["row"]}: {e["error"]}')
