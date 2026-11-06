"""Add OOXML features unavailable/reliably unexported by the workbook authoring engine.
Power Query format: Microsoft MS-QDEFF; PivotTable: SpreadsheetML.
No Excel engine is invoked; refresh must be verified in desktop Excel.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
import json,struct,io,base64,uuid,argparse

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
C='http://schemas.openxmlformats.org/package/2006/content-types'
def element(tag,ns=S,**attrs):
    return E.Element('{'+ns+'}'+tag,**{k:str(v) for k,v in attrs.items()})
def sub(parent,tag,ns=S,**attrs):
    x=element(tag,ns,**attrs);parent.append(x);return x
def xml(x):return E.tostring(x,encoding='UTF-8',xml_declaration=True,standalone=True)
def part(b):return struct.pack('<I',len(b))+b

def add(root):
    p=root/'10_Расчёт_вознаграждений.xlsx'
    d=json.loads((root/'qa/analysis.json').read_text(encoding='utf-8'))
    with ZipFile(p) as z: parts={n:z.read(n) for n in z.namelist()}
    types=E.fromstring(parts['[Content_Types].xml'])
    def override(path,typ):
        sub(types,'Override',C,PartName='/'+path,ContentType=typ)
    def rel(path,typ,target,rid):
        tree=E.fromstring(parts[path]) if path in parts else element('Relationships',P)
        sub(tree,'Relationship',P,Id=rid,Type=R+'/'+typ,Target=target)
        parts[path]=xml(tree)
    # Portable workbook named input, not a hard-coded machine path in M.
    book=E.fromstring(parts['xl/workbook.xml'])
    defined=book.find('{'+S+'}definedNames')
    if defined is None:
        defined=element('definedNames')
        calc=book.find('{'+S+'}calcPr')
        book.insert(list(book).index(calc) if calc is not None else len(book),defined)
    sub(defined,'definedName',name='SourceFolder').text="'Параметры'!$B$2"
    queries=[('Услуги','ServicesData','Услуги',5),('Уникальные','UniquePatients','Уникальные',6),
             ('ПациентМесяц','PatientMonths','Пациент месяц',4),('Вознаграждения','PartnerRewards','Вознаграждения',3),
             ('ИтогиМесяцев','MonthlySummary','Итоги',1),('ИтогиПериода','PeriodSummary','Итоги',1)]
    # Build query package and per-query load metadata from public schema.
    pack=io.BytesIO()
    with ZipFile(pack,'w',ZIP_DEFLATED) as z:
        z.writestr('Config/Package.xml','<?xml version="1.0" encoding="utf-8"?><Package><Version>2.132.328.0</Version><MinVersion>2.21.0.0</MinVersion><Culture>ru-RU</Culture></Package>')
        z.writestr('[Content_Types].xml',f'<Types xmlns="{C}"><Default Extension="xml" ContentType="text/xml"/><Default Extension="m" ContentType="application/x-ms-m"/></Types>')
        mfile=root/'power_query/Section1.m'
        if not mfile.exists():mfile=root/'Section1.m'
        z.writestr('Formulas/Section1.m',mfile.read_bytes())
    metadata=E.Element('LocalPackageMetadataFile')
    items=E.SubElement(metadata,'Items')
    for q,t,s,si in queries:
        item=E.SubElement(items,'Item');loc=E.SubElement(item,'ItemLocation')
        E.SubElement(loc,'ItemType').text='Formula';E.SubElement(loc,'ItemPath').text='Section1/'+q
        entries=E.SubElement(item,'StableEntries')
        for k,v in {'FillEnabled':'l1','FillToDataModelEnabled':'l0','FillTarget':'s'+t,
                    'RecoveryTargetSheet':'s'+s,'RecoveryTargetRow':'l'+('3' if s=='Вознаграждения' else '4' if t=='MonthlySummary' else '9' if t=='PeriodSummary' else '0'),
                    'RecoveryTargetColumn':'l0','ResultType':'sTable','IsPrivate':'l0'}.items():
            E.SubElement(entries,'Entry',Type=k,Value=v)
    binary=io.BytesIO()
    with ZipFile(binary,'w'):pass
    meta=struct.pack('<I',0)+part(xml(metadata))+part(binary.getvalue())
    permissions=b'<PermissionList><CanEvaluateFuturePackages>false</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>'
    mashup=struct.pack('<I',0)+part(pack.getvalue())+part(permissions)+part(meta)+part(b'')
    dm=E.Element('{http://schemas.microsoft.com/DataMashup}DataMashup',nsmap={None:'http://schemas.microsoft.com/DataMashup'},sqmid='{'+str(uuid.uuid4()).upper()+'}')
    dm.text=base64.b64encode(mashup).decode()
    parts['customXml/item1.xml']=xml(dm)
    props=E.Element('{http://schemas.openxmlformats.org/officeDocument/2006/customXml}datastoreItem',nsmap={'ds':'http://schemas.openxmlformats.org/officeDocument/2006/customXml'})
    props.set('{http://schemas.openxmlformats.org/officeDocument/2006/customXml}itemID','{'+str(uuid.uuid4()).upper()+'}')
    refs=E.SubElement(props,'{http://schemas.openxmlformats.org/officeDocument/2006/customXml}schemaRefs')
    ref=E.SubElement(refs,'{http://schemas.openxmlformats.org/officeDocument/2006/customXml}schemaRef');ref.set('{http://schemas.openxmlformats.org/officeDocument/2006/customXml}uri','http://schemas.microsoft.com/DataMashup')
    parts['customXml/itemProps1.xml']=xml(props)
    rel('customXml/_rels/item1.xml.rels','customXmlProps','itemProps1.xml','rId1')
    rel('xl/_rels/workbook.xml.rels','customXml','../customXml/item1.xml','rIdMashup')
    override('customXml/itemProps1.xml','application/vnd.openxmlformats-officedocument.customXmlProperties+xml')
    # QueryTable relationships attach native Power Query results to existing Excel tables.
    connections=element('connections')
    tables={}
    for name in list(parts):
        if name.startswith('xl/tables/table') and name.endswith('.xml'):
            t=E.fromstring(parts[name]);tables[t.get('name')]=(name,t)
    for index,(q,t,s,si) in enumerate(queries,1):
        conn=sub(connections,'connection',id=index,name='Запрос — '+q,type=5,refreshedVersion=8,background=0,saveData=1)
        sub(conn,'dbPr',connection=f'Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={q};Extended Properties=""',command=f'SELECT * FROM [{q}]')
        tablepath,tree=tables[t];tree.set('tableType','queryTable')
        cols=tree.find('{'+S+'}tableColumns')
        qt=element('queryTable',name=t,connectionId=index,autoFormatId=16,applyNumberFormats=0,applyBorderFormats=0,applyFontFormats=0,applyPatternFormats=0,applyAlignmentFormats=0,applyWidthHeightFormats=1)
        refresh=sub(qt,'queryTableRefresh',nextId=len(cols)+1)
        fields=sub(refresh,'queryTableFields',count=len(cols))
        for j,col in enumerate(cols,1):
            col.set('queryTableFieldId',str(j))
            sub(fields,'queryTableField',id=j,name=col.get('name'),tableColumnId=col.get('id'))
        parts[tablepath]=xml(tree)
        qpath=f'xl/queryTables/queryTable{index}.xml';parts[qpath]=xml(qt)
        rel('xl/tables/_rels/'+Path(tablepath).name+'.rels','queryTable',f'../queryTables/queryTable{index}.xml','rIdQueryTable')
        override(qpath,'application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml')
    parts['xl/connections.xml']=xml(connections)
    rel('xl/_rels/workbook.xml.rels','connections','connections.xml','rIdConnections')
    override('xl/connections.xml','application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml')
    # Distinct patient keys in the source table allow a normal COUNT PivotTable.
    headers=['Месяц','клиника услуги','Канал привлечения','id пациента','Единица']
    unique=d['unique']; shared=[sorted({str(r[i]) for r in unique}) for i in range(4)]+[]
    cache=element('pivotCacheDefinition',recordCount=len(unique),createdVersion=6,refreshedVersion=6,minRefreshableVersion=3,saveData=1,refreshOnLoad=0)
    cache.set('{'+R+'}id','rIdRecords')
    source=sub(cache,'cacheSource',type='worksheet');sub(source,'worksheetSource',name='UniquePatients')
    cf=sub(cache,'cacheFields',count=5)
    for i,h in enumerate(headers):
        field=sub(cf,'cacheField',name=h,numFmtId=0)
        if i<4:
            ss=sub(field,'sharedItems',count=len(shared[i]))
            for v in shared[i]:sub(ss,'s',v=v)
        else:sub(field,'sharedItems',containsString=0,containsNumber=1,containsInteger=1,containsSemiMixedTypes=0)
    records=element('pivotCacheRecords',count=len(unique))
    for row in unique:
        rec=sub(records,'r')
        for i,v in enumerate(row):sub(rec,'x' if i<4 else 'n',v=shared[i].index(str(v)) if i<4 else v)
    parts['xl/pivotCache/pivotCacheDefinition1.xml']=xml(cache)
    parts['xl/pivotCache/pivotCacheRecords1.xml']=xml(records)
    rel('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels','pivotCacheRecords','pivotCacheRecords1.xml','rIdRecords')
    rel('xl/_rels/workbook.xml.rels','pivotCacheDefinition','pivotCache/pivotCacheDefinition1.xml','rIdPivotCache')
    caches=element('pivotCaches');pc=sub(caches,'pivotCache',cacheId=1);pc.set('{'+R+'}id','rIdPivotCache')
    calc=book.find('{'+S+'}calcPr');book.insert(list(book).index(calc) if calc is not None else len(book),caches)
    pairs=sorted({(r[1],r[2]) for r in unique})
    pivot=element('pivotTableDefinition',name='УникальныеПациенты',cacheId=1,dataCaption='Уникальные пациенты',rowGrandTotals=0,colGrandTotals=0,compact=0,compactData=0,outline=0,outlineData=0,gridDropZones=1,updatedVersion=6,minRefreshableVersion=3,createdVersion=6)
    sub(pivot,'location',ref=f'A4:E{len(pairs)+4}',firstHeaderRow=0,firstDataRow=1,firstDataCol=2)
    fields=sub(pivot,'pivotFields',count=5)
    for i in range(5):
        attrs={'showAll':0}
        if i==0:attrs['axis']='axisCol'
        if i in (1,2):attrs.update(axis='axisRow',compact=0,outline=0,defaultSubtotal=0)
        if i==3:attrs['dataField']=1
        field=sub(fields,'pivotField',**attrs)
        if i<3:
            its=sub(field,'items',count=len(shared[i]))
            for j in range(len(shared[i])):sub(its,'item',x=j)
    rf=sub(pivot,'rowFields',count=2);sub(rf,'field',x=1);sub(rf,'field',x=2)
    ri=sub(pivot,'rowItems',count=len(pairs))
    for clinic,channel in pairs:
        item=sub(ri,'i');sub(item,'x',v=shared[1].index(clinic));sub(item,'x',v=shared[2].index(channel))
    cfs=sub(pivot,'colFields',count=1);sub(cfs,'field',x=0)
    ci=sub(pivot,'colItems',count=len(shared[0]))
    for j in range(len(shared[0])):sub(sub(ci,'i'),'x',v=j)
    df=sub(pivot,'dataFields',count=1);sub(df,'dataField',name='Уникальных пациентов',fld=3,subtotal='count',baseField=0,baseItem=0)
    sub(pivot,'pivotTableStyleInfo',name='PivotStyleMedium9',showRowHeaders=1,showColHeaders=1,showRowStripes=1,showColStripes=0,showLastColumn=1)
    parts['xl/pivotTables/pivotTable1.xml']=xml(pivot)
    rel('xl/pivotTables/_rels/pivotTable1.xml.rels','pivotCacheDefinition','../pivotCache/pivotCacheDefinition1.xml','rIdCache')
    rel('xl/worksheets/_rels/sheet2.xml.rels','pivotTable','../pivotTables/pivotTable1.xml','rIdPivot')
    sheet=E.fromstring(parts['xl/worksheets/sheet2.xml'])
    pp=sub(sheet,'pivotTableParts',count=1);piv=sub(pp,'pivotTablePart');piv.set('{'+R+'}id','rIdPivot')
    parts['xl/worksheets/sheet2.xml']=xml(sheet)
    for fn,ct in [('pivotCache/pivotCacheDefinition1.xml','pivotCacheDefinition'),('pivotCache/pivotCacheRecords1.xml','pivotCacheRecords'),('pivotTables/pivotTable1.xml','pivotTable')]:
        override('xl/'+fn,'application/vnd.openxmlformats-officedocument.spreadsheetml.'+ct+'+xml')
    parts['xl/workbook.xml']=xml(book);parts['[Content_Types].xml']=xml(types)
    # XML integrity and duplicate-part guards before replacing the file.
    for name,data in parts.items():
        if name.endswith('.xml') or name.endswith('.rels'):E.fromstring(data)
    with ZipFile(p,'w',ZIP_DEFLATED) as z:
        for name,data in parts.items():z.writestr(name,data)
    print('Native query connections:',len(queries),'Pivot records:',len(unique))

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('root',type=Path);add(parser.parse_args().root)
