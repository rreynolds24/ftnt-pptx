#!/usr/bin/env python3
"""Inspect/export supplied brand artwork; import local Draw.io libraries safely.
No network calls. Python 3.10+ standard library for inspect/import. PDF crops
require the optional local PyMuPDF dependency and a hash-pinned source recipe.
Imported artwork is staged as UNVERIFIED until provenance and visual review pass.
"""
from __future__ import annotations
import argparse,base64,hashlib,json,re,sys
from pathlib import Path
from urllib.parse import unquote_to_bytes
from xml.dom import minidom
ROOT=Path(__file__).resolve().parent.parent

def sha(b):return hashlib.sha256(b).hexdigest()
def parse(b):
    if b'<!DOCTYPE' in b.upper() or b'<!ENTITY' in b.upper():raise ValueError('DTD/entities not allowed')
    return minidom.parseString(b)
def safe_svg(b):
    d=parse(b)
    if d.documentElement.localName!='svg':raise ValueError('Expected SVG')
    for n in d.getElementsByTagName('*'):
        if n.localName.lower() in ('script','foreignobject','iframe','object','embed','image','feimage'):
            raise ValueError('Active or nested-image SVG is not imported by this conservative importer')
        for k in list(n.attributes.keys()):
            value=n.getAttribute(k).strip()
            if k.lower().startswith('on'):raise ValueError('SVG event handler rejected')
            if k.lower().endswith('href') and value and not value.startswith('#'):raise ValueError('External SVG link rejected')
        if n.localName=='style' or n.hasAttribute('style'):
            css=n.getAttribute('style')+' '+''.join(c.data for c in n.childNodes if c.nodeType in (c.TEXT_NODE,c.CDATA_SECTION_NODE))
            if re.search(r'@import|@font-face|url\s*\(\s*[\'"]?(?!#)',css,re.I):raise ValueError('External CSS/font references rejected')
    return d

def import_library(path,out,source_label):
    if out.exists():raise ValueError('Import output must be a new directory')
    raw=path.read_bytes()
    if len(raw)>64*1024*1024:raise ValueError('Library is too large')
    d=parse(raw)
    if d.documentElement.tagName!='mxlibrary':raise ValueError('Expected local Draw.io <mxlibrary> XML')
    items=json.loads(''.join(c.data for c in d.documentElement.childNodes if c.nodeType in (c.TEXT_NODE,c.CDATA_SECTION_NODE)))
    if not isinstance(items,list) or len(items)>5000:raise ValueError('Invalid or excessive library entries')
    staged=[];skipped=[];total=0
    for i,item in enumerate(items):
        try:
            uri=item.get('data','');head,body=uri.split(',',1)
            if not head.startswith('data:image/'):raise ValueError('Only embedded image data URIs supported, not compressed shape XML')
            rawimg=base64.b64decode(body,validate=True) if ';base64' in head else unquote_to_bytes(body)
            total+=len(rawimg)
            if len(rawimg)>8*1024*1024 or total>128*1024*1024:raise ValueError('Image resource limit')
            if head.startswith('data:image/svg+xml'):safe_svg(rawimg);ext='.svg'
            elif head.startswith('data:image/png') and rawimg[:8]==b'\x89PNG\r\n\x1a\n':ext='.png'
            else:raise ValueError('Unsupported embedded format')
            title=str(item.get('title') or 'unnamed')
            slug=re.sub('[^a-z0-9]+','-',title.casefold()).strip('-')[:70] or 'asset'
            name=slug+'-'+sha(rawimg)[:10]+ext
            staged.append((name,rawimg,{'id':slug+'-'+sha(rawimg)[:10],'label':title,'path':name,'sha256':sha(rawimg),'kind':'icon','width_px':item.get('w'),'height_px':item.get('h'),'approved':False,'provenance':{'origin':'local-drawio-library','label':source_label,'library_sha256':sha(raw),'index':i},'review_required':'Check provider authority, identity, original colours, render and size. Add PNG fallback/ink bounds before admission to production catalog.'}))
        except (ValueError,KeyError,TypeError) as exc:skipped.append({'index':i,'reason':str(exc)})
    out.mkdir(parents=True)
    for name,data,entry in staged:(out/name).write_bytes(data)
    manifest={'schema_version':1,'assets':[x[2] for x in staged],'skipped':skipped,'network_used':False}
    (out/'asset-catalog.staged.json').write_text(json.dumps(manifest,indent=2)+'\n')
    return {'imported':len(staged),'skipped':len(skipped),'review_required':True,'output':str(out)}

def extract_pdf(path,recipe_path,out):
    import fitz
    recipe=json.loads(recipe_path.read_text())
    if sha(path.read_bytes())!=recipe['source_sha256']:raise ValueError('Source PDF does not match crop recipe hash')
    if out.exists():raise ValueError('Extraction output must be a new directory')
    d=fitz.open(path);out.mkdir(parents=True);entries=[]
    for c in recipe['crops']:
        page=d[c['page_1_based']-1];rect=fitz.Rect(c['crop_points'])
        if not page.rect.contains(rect):raise ValueError('Crop is outside page')
        # Render the exact crop; do not export the whole page disguised as tiny SVG.
        scale=c.get('pixels_per_point',12)
        pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale),clip=rect,alpha=True)
        # Optional source-specific exclusion removes adjacent diagram content,
        # never part of the icon. Its nonintersection with artwork needs visual QA.
        for excluded in c.get('exclude_regions_points',[]):
            region=fitz.Rect(excluded)
            if not rect.contains(region):raise ValueError('Exclusion is outside crop')
            pix.set_rect(fitz.IRect(region.x0*scale,region.y0*scale,region.x1*scale,region.y1*scale),(0,0,0,0))
        if not re.fullmatch('[a-z0-9-]+',c['id']):raise ValueError('Unsafe crop ID')
        name=c['id']+'.png';pix.save(out/name)
        entries.append({'id':c['id'],'path':name,'sha256':sha((out/name).read_bytes()),'source_sha256':recipe['source_sha256'],'page_1_based':c['page_1_based'],'crop_points':c['crop_points'],'exclude_regions_points':c.get('exclude_regions_points',[]),'treatment':'raster crop; any listed exclusions remove adjacent diagram content only','approved':False})
    (out/'crop-manifest.json').write_text(json.dumps(entries,indent=2)+'\n')
    return {'extracted':len(entries),'output':str(out),'visual_review_required':True}

def main():
    p=argparse.ArgumentParser(description=__doc__);s=p.add_subparsers(dest='cmd',required=True)
    i=s.add_parser('import-drawio');i.add_argument('library',type=Path);i.add_argument('--output',type=Path,required=True);i.add_argument('--source-label',required=True)
    c=s.add_parser('extract-pdf');c.add_argument('pdf',type=Path);c.add_argument('--recipe',type=Path,required=True);c.add_argument('--output',type=Path,required=True)
    a=p.parse_args()
    try:
        r=import_library(a.library,a.output,a.source_label) if a.cmd=='import-drawio' else extract_pdf(a.pdf,a.recipe,a.output)
        print(json.dumps(r));return 0
    except Exception as exc:print('ERROR: '+str(exc),file=sys.stderr);return 2
if __name__=='__main__':sys.exit(main())
