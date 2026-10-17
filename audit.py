"""Independent reconciliation of the saved XLSX, native parts and source exports."""
from pathlib import Path
from collections import defaultdict
from decimal import Decimal
from zipfile import ZipFile
import json,base64,struct,io,argparse
import openpyxl
from lxml import etree
from analyze import analyze,verify,number

def audit(root,source):
    verify();d=analyze(source)
    path=root/'Клиника_готовое_тестовое.xlsx'
    w=openpyxl.load_workbook(path,data_only=True)
    f=openpyxl.load_workbook(path,data_only=False)
    assert w['ЕКМС'].max_row-1==d['controls']['rows']
    assert w['Уникальные'].max_row-1==len(d['unique'])
    assert w['Пациент месяц'].max_row-1==len(d['ledger'])
    assert sum(number(w['ЕКМС'].cell(i,8).value) for i in range(2,w['ЕКМС'].max_row+1))==number(d['controls']['compensation'])
    ledger={(r[0],r[1]):r for r in d['ledger']}
    for row in list(w['Пациент месяц'].values)[1:]:
        expected=ledger[(row[0],str(row[1]))]
        assert number(row[4])==number(expected[4]) and row[5]==expected[5]
    assert number(w['Итоги']['B14'].value)==number(d['controls']['reward'])
    expected=defaultdict(int)
    for m,c,ch,pid,one in d['unique']:expected[(m,c,ch)]+=1
    for row in list(w['Сводная'].values)[4:]:
        for j,m in enumerate(d['months'],2):assert row[j]==expected[(m,row[0],row[1])]
    assert len(f['Сводная']._pivots)==1
    assert f['Сводная']._pivots[0].cache.recordCount==len(d['unique'])
    errors=[]
    for s in w:
        for row in s:
            for c in row:
                if c.data_type=='e':errors.append((s.title,c.coordinate,c.value))
    assert not errors,errors
    with ZipFile(path) as z:
        assert len(z.namelist())==len(set(z.namelist()))
        for name in z.namelist():
            if name.endswith(('.xml','.rels')):etree.fromstring(z.read(name))
        connections=etree.fromstring(z.read('xl/connections.xml'));assert len(connections)==6
        dm=etree.fromstring(z.read('customXml/item1.xml'))
        b=base64.b64decode(dm.text);n=struct.unpack('<I',b[4:8])[0]
        with ZipFile(io.BytesIO(b[8:8+n])) as m:
            assert m.read('Formulas/Section1.m')==(root/'power_query/Section1.m' if (root/'power_query/Section1.m').exists() else root/'Section1.m').read_bytes()
        # Every internal package relationship points to an existing part.
        import posixpath
        for name in z.namelist():
            if not name.endswith('.rels'):continue
            folder=posixpath.dirname(posixpath.dirname(name))
            for rel in etree.fromstring(z.read(name)):
                if rel.get('TargetMode')=='External':continue
                target=posixpath.normpath(posixpath.join(folder,rel.get('Target'))).lstrip('/')
                assert target in z.namelist(),(name,target)
    report={'status':'PASS','controls':d['controls'],'monthly':d['summary'],
            'checks':['source rows and compensation','patient-month payments and payouts','distinct patient counts per group','formula caches without Excel errors','native pivot cache','six Power Query connections','embedded M matches source','XML and relationship integrity','threshold and age boundary cases'],
            'desktop_excel_refresh':'NOT RUN — Excel is not installed in this environment'}
    (root/'qa/audit.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('root',type=Path);p.add_argument('source',type=Path);a=p.parse_args();audit(a.root,a.source)
