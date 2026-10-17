"""Read EKMS exports; validate and prepare auditable workbook inputs. No patient names required."""
from pathlib import Path
from decimal import Decimal
from datetime import datetime
from collections import defaultdict
import argparse, json
import openpyxl

def number(v):
    if v is None or str(v).strip() == '':
        raise ValueError('Missing required number')
    return Decimal(str(v).replace('\xa0', '').replace(' ', '').replace(',', '.'))

def date(v):
    return v if isinstance(v, datetime) else datetime.strptime(str(v)[:10], '%d.%m.%Y')

def clean(v):
    return ' '.join(str(v).split()) if v is not None else ''

def reward(month, child, amount, representative):
    if not representative:
        return 0
    if month == '2024-07':
        return 5000 if child and amount > 100000 else 0
    if month in ('2024-08', '2024-09'):
        return 8000 if amount > 120000 else 0
    raise ValueError('No reward policy for month ' + month)

def analyze(folder):
    rows, files = [], []
    for p in sorted(Path(folder).glob('EKMS_*.xlsx')):
        ws = openpyxl.load_workbook(p, data_only=True)['EKMS']
        raw = list(ws.values)
        headers = [clean(x).lower() for x in raw[0]]
        count = 0
        for lineno, values in enumerate(raw[1:], 2):
            if all(x is None for x in values):
                continue
            r = dict(zip(headers, values))
            pid = clean(r['id пациента'])
            if not pid:
                raise ValueError(f'{p.name}:{lineno}: missing patient ID')
            age, amount = number(r['возраст']), number(r['компенсация, руб.'])
            if age < 0 or age > 120:
                raise ValueError('Invalid age')
            rep, clinic = clean(r['представитель']), clean(r['клиника услуги'])
            if not clinic:
                raise ValueError('Missing clinic')
            d = date(r['дата услуги'])
            month = d.strftime('%Y-%m')
            rows.append([p.name, lineno, pid, float(age), d.strftime('%Y-%m-%d'), month,
                         clinic, float(amount), rep, 'дети' if age < 18 else 'взрослые',
                         'от Агентов' if rep else 'Реклама'])
            count += 1
        files.append([p.name, count])
    if not rows:
        raise ValueError('No input records')
    groups = defaultdict(list)
    for r in rows:
        groups[(r[5], r[2])].append(r)
    ledger = []
    for (month, pid), records in sorted(groups.items()):
        reps = {r[8] for r in records if r[8]}
        cats = {r[9] for r in records}
        if len(reps) > 1 or len(cats) > 1:
            raise ValueError('Ambiguous representative or age category for a patient-month')
        rep = next(iter(reps), '')
        total = sum(number(r[7]) for r in records)
        child = records[0][9] == 'дети'
        ledger.append([month, pid, rep, records[0][9], float(total),
                       reward(month, child, total, rep), len(records)])
    unique = sorted({(r[5], r[6], r[10], r[2]) for r in rows})
    counts = defaultdict(int)
    for month, clinic, channel, pid in unique:
        counts[(month, clinic, channel)] += 1
    reps = sorted({r[2] for r in ledger if r[2]})
    months = sorted({r[5] for r in rows})
    payments = [[rep] + [sum(r[5] for r in ledger if r[2] == rep and r[0] == m) for m in months] for rep in reps]
    summary = [[m, sum(r[5] == m for r in rows), len({r[2] for r in rows if r[5] == m}),
                float(sum(number(r[7]) for r in rows if r[5] == m)),
                sum(r[5] > 0 for r in ledger if r[0] == m), sum(r[5] for r in ledger if r[0] == m)] for m in months]
    return dict(rows=rows, ledger=ledger, unique=[list(x)+[1] for x in unique],
                counts=[list(k)+[v] for k,v in sorted(counts.items())], payments=payments,
                summary=summary, files=files, months=months,
                controls={'rows':len(rows),'patients_period':len({r[2] for r in rows}),
                          'compensation':float(sum(number(r[7]) for r in rows)),
                          'reward':sum(r[5] for r in ledger),'patient_months':len(ledger)})

def verify():
    for amount in (99999.99,100000,100000.01):
        assert reward('2024-07', True, Decimal(str(amount)), 'Partner') == (5000 if amount>100000 else 0)
    assert reward('2024-07',False,Decimal(200000),'Partner') == 0
    assert reward('2024-08',False,Decimal(120000),'Partner') == 0
    assert reward('2024-08',False,Decimal('120000.01'),'Partner') == 8000
    assert reward('2024-09',True,Decimal(200000),'') == 0
    assert number('17,99') < 18 and number('18,00') >= 18
    assert clean('  \xa0 ') == ''
    assert reward('2024-08',False,Decimal(70000)+Decimal(60000),'Partner') == 8000

if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('folder',type=Path)
    parser.add_argument('--out',type=Path,default=Path('analysis.json'))
    args=parser.parse_args()
    verify()
    result=analyze(args.folder)
    args.out.parent.mkdir(parents=True,exist_ok=True)
    args.out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'controls':result['controls'],'months':result['summary']},ensure_ascii=False))
