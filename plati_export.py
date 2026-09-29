"""Exportă sumele din foaia „Evidență plăți" pentru pagina plati.html.

Rulare:  python plati_export.py
Citește cel mai recent __SAH 2627__pilot*.xlsx din OneDrive (pilot, pilot_v1, ...),
dintr-o copie, ca să meargă și cu Excel deschis,
și scrie plati_date.js. Numele copiilor NU apar în clar în fișier: fiecare copil e găsit
după amprenta SHA-256 a numelui normalizat (fără diacritice, cuvintele în ordine alfabetică).
"""
import glob, hashlib, itertools, json, shutil, tempfile, unicodedata, datetime, os

FOLDER = os.path.expanduser(r'~\OneDrive\__SAH_26-27__')
TIPAR = '__SAH 2627__pilot*.xlsx'   # fișierele din „arhiva versiuni vechi” nu intră (subfolder)
DEST = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'plati_date.js')
SARE = 'lab-sah-2627'   # trebuie să fie identică cu cea din plati.html

LUNI = ['septembrie', 'octombrie', 'noiembrie', 'decembrie', 'ianuarie',
        'februarie', 'martie', 'aprilie', 'mai', 'iunie']
PRIMA_COL = 4           # D = Prez. septembrie; fiecare lună are 3 coloane (Prez., De plată, Achitat)
PRIMUL_RAND = 7


def normalizeaza(text):
    text = unicodedata.normalize('NFD', str(text)).encode('ascii', 'ignore').decode()
    text = ''.join(c if c.isalpha() else ' ' for c in text.upper())
    return text.split()


def chei(nume):
    """Amprente pentru numele complet și pentru orice combinație de cel puțin 2 cuvinte
    (ca „Iacovici Iris" să-l găsească și pe „IACOVICI IRIS MIHAELA")."""
    cuvinte = normalizeaza(nume)
    rez = set()
    for k in range(min(2, len(cuvinte)), len(cuvinte) + 1):
        for comb in itertools.combinations(cuvinte, k):
            rez.add(hashlib.sha256((SARE + ' '.join(sorted(comb))).encode()).hexdigest()[:20])
    return rez


def numar(v):
    return v if isinstance(v, (int, float)) else 0


def main():
    import openpyxl
    fisiere = [f for f in glob.glob(os.path.join(FOLDER, TIPAR))
               if not os.path.basename(f).startswith('~$')]   # ~$... = fișierul de blocare al Excel
    SURSA = max(fisiere, key=os.path.getmtime)
    print('Sursa:', os.path.basename(SURSA))
    copie = os.path.join(tempfile.gettempdir(), 'plati_export_copie.xlsx')
    try:
        shutil.copyfile(SURSA, copie)
    except PermissionError:   # Excel ține fișierul blocat; Copy-Item din PowerShell reușește
        import subprocess
        subprocess.run(['powershell', '-NoProfile', '-Command',
                        f'Copy-Item -LiteralPath "{SURSA}" -Destination "{copie}" -Force'], check=True)
    ws = openpyxl.load_workbook(copie, data_only=True)['Evidență plăți']
    tarif = ws['C3'].value

    copii, index = [], {}
    for r in range(PRIMUL_RAND, ws.max_row + 1):
        nume = ws.cell(r, 2).value
        if not nume or str(nume).strip().upper() == 'TOTAL':
            continue
        luni = []
        for i in range(len(LUNI)):
            c = PRIMA_COL + 3 * i
            luni.append([numar(ws.cell(r, c).value), numar(ws.cell(r, c + 1).value),
                         numar(ws.cell(r, c + 2).value)])
        idx = len(copii)
        copii.append(luni)
        for k in chei(nume):
            index.setdefault(k, []).append(idx)

    # luna curentă = ultima lună în care a venit cineva
    curenta = max((i for i in range(len(LUNI)) if any(c[i][0] for c in copii)), default=0)

    date = {
        'actualizat': datetime.date.today().strftime('%d.%m.%Y'),
        'tarif': tarif,
        'luni': LUNI,
        'lunaCurenta': curenta,
        'copii': copii,
        'index': index,
    }
    with open(DEST, 'w', encoding='utf-8') as f:
        f.write('// Generat de plati_export.py — nu modifica de mână.\n')
        f.write('window.PLATI = ' + json.dumps(date, ensure_ascii=False, separators=(',', ':')) + ';\n')
    print(f'{len(copii)} copii exportați, luna curentă: {LUNI[curenta]}, tarif {tarif} lei -> {DEST}')


if __name__ == '__main__':
    main()
