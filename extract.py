"""Extract in-sheet images from a Google Sheets xlsx export, mapped to the row/column they sit in.
Usage: python3 extract.py <file.xlsx> <out_dir>   -> writes images + manifest.json"""
import sys, zipfile, re, json, os, posixpath
import xml.etree.ElementTree as ET
NS = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'm': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
RID = '{%s}id' % NS['r']
EMBED = '{%s}embed' % NS['r']

def rels(z, path):
    d, b = posixpath.split(path)
    rp = posixpath.join(d, '_rels', b + '.rels')
    if rp not in z.namelist(): return {}
    return {r.get('Id'): posixpath.normpath(posixpath.join(d, r.get('Target')))
            for r in ET.fromstring(z.read(rp)).findall('rel:Relationship', NS)
            if r.get('TargetMode') != 'External'}

def main(xlsx, out):
    os.makedirs(out, exist_ok=True)
    z = zipfile.ZipFile(xlsx)
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    wbrels = rels(z, 'xl/workbook.xml')
    manifest = []
    for s in wb.find('m:sheets', NS):
        title, sheet_path = s.get('name'), wbrels[s.get(RID)]
        sx = ET.fromstring(z.read(sheet_path))
        drw = sx.find('m:drawing', NS)
        if drw is None: continue
        dpath = rels(z, sheet_path)[drw.get(RID)]
        drels = rels(z, dpath)
        for anchor in ET.fromstring(z.read(dpath)):
            frm = anchor.find('xdr:from', NS)
            blip = anchor.find('.//a:blip', NS)
            if frm is None or blip is None: continue
            row = int(frm.find('xdr:row', NS).text) + 1   # 1-indexed Excel row
            col = int(frm.find('xdr:col', NS).text) + 1
            media = drels[blip.get(EMBED)]
            manifest.append({'sheet': title, 'row': row, 'col': col, 'media': media})
    for m in manifest:
        dest = os.path.join(out, posixpath.basename(m['media']))
        if not os.path.exists(dest):
            open(dest, 'wb').write(z.read(m['media']))
        m['file'] = dest
    json.dump(manifest, open(os.path.join(out, 'manifest.json'), 'w'), indent=1)
    print(f'{len(manifest)} anchored images -> {out}')

if __name__ == '__main__':
    main(*sys.argv[1:3])
