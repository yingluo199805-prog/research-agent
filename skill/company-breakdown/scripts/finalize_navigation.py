#!/usr/bin/env python3
"""Add internal navigation without reserializing formulas, caches or native Excel formatting.
Usage: python3 finalize_navigation.py workbook_spec.json output.xlsx
The artifact renderer does not currently provide an exported hyperlink API.
"""
import json, pathlib, sys, tempfile, zipfile, xml.etree.ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)

def finalize(spec_path, output):
    spec=json.loads(pathlib.Path(spec_path).read_text(encoding='utf-8'))
    p=pathlib.Path(output);links={};names={s['name']:s for s in spec['sheets']}
    for i,s in enumerate(spec['sheets']):
        entries=list(s.get('links',[]))
        if s.get('layout')=='revenue':
            if not any(x['cell']=='A1' for x in entries):entries.append({'cell':'A1','location':"'"+s['name'].replace("'","''")+"'!C3"})
            if not any(x['cell']=='A2' for x in entries):entries.append({'cell':'A2','location':"'"+s['name'].replace("'","''")+"'!C37"})
        if s['name']=='目录':
            for r,row in enumerate(s['rows'],1):
                for c,value in enumerate(row,1):
                    if isinstance(value,str) and value in names and value!='目录':
                        label='';k=c
                        while k:k,n=divmod(k-1,26);label=chr(65+n)+label
                        entries.append({'cell':label+str(r),'location':"'"+value.replace("'","''")+"'!A1"})
        if entries:links[i]=entries
    if not links:return
    after={'sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations'}
    with tempfile.NamedTemporaryFile(dir=p.parent,suffix='.xlsx',delete=False) as f:tmp=pathlib.Path(f.name)
    try:
        with zipfile.ZipFile(p) as zin,zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data=zin.read(info.filename)
                for i,entries in links.items():
                    if info.filename!=f'xl/worksheets/sheet{i+1}.xml':continue
                    root=ET.fromstring(data);hl=root.find('{'+NS+'}hyperlinks')
                    if hl is None:
                        hl=ET.Element('{'+NS+'}hyperlinks');at=0
                        for j,child in enumerate(root):
                            if child.tag.split('}')[-1] in after:at=j+1
                        root.insert(at,hl)
                    for entry in entries:
                        for existing in list(hl):
                            if existing.attrib.get('ref')==entry['cell']:hl.remove(existing)
                        ET.SubElement(hl,'{'+NS+'}hyperlink',ref=entry['cell'],location=entry['location'])
                    data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
                zout.writestr(info,data)
        tmp.replace(p)
    finally:
        if tmp.exists():tmp.unlink()

if __name__=='__main__':
    if len(sys.argv)!=3:raise SystemExit(__doc__)
    finalize(sys.argv[1],sys.argv[2])
