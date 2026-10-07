#!/usr/bin/env python3
"""Offline, non-destructive Fortinet PPTX audit/plan/apply utility.

Python 3.10+ standard library only. No network, PowerPoint automation, or model calls.
Preserves textual content and untouched ZIP parts; edits only reviewed formatting.
See references/offline-pptx-toolkit.md for supported scope and limitations.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
import math
import os
from pathlib import Path, PurePosixPath
import posixpath
import re
import sys
import tempfile
import zipfile
from xml.dom import minidom, Node

VERSION = '2.0.0'
ROOT = Path(__file__).resolve().parent.parent
EMU = 914400
NS = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'rel':'http://schemas.openxmlformats.org/package/2006/relationships',
      'ct':'http://schemas.openxmlformats.org/package/2006/content-types'}
SYMBOL_FONTS = ('wingdings','webdings','symbol','font awesome','fontawesome','material icons')
MAX_ENTRIES, MAX_ENTRY, MAX_TOTAL = 10000, 64*1024*1024, 512*1024*1024
STYLE_PART = re.compile(r'^ppt/(?:(?:slides|slideMasters|slideLayouts|notesSlides|notesMasters|charts|diagrams|theme)/[^/]+|presentation|tableStyles)\.xml$')
HEX = re.compile(r'^[0-9A-Fa-f]{6}$')

class BrandError(ValueError):
    """A fail-closed input, scope, or integrity error."""

def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def jd(path: Path) -> dict:
    return json.loads(path.read_text(encoding='utf-8'))

def write_json(path: Path, value: dict, force: bool = False) -> None:
    if path.exists() and not force:
        raise BrandError(f'Refusing to overwrite {path}')
    path.parent.mkdir(parents=True, exist_ok=True)
    data=(json.dumps(value,indent=2,ensure_ascii=False)+'\n').encode()
    atomic(path,data)

def atomic(path: Path, data: bytes) -> None:
    fd,tmp=tempfile.mkstemp(prefix='.fortinet-',suffix='.tmp',dir=path.parent)
    try:
        with os.fdopen(fd,'wb') as f:
            f.write(data); f.flush(); os.fsync(f.fileno())
        os.replace(tmp,path)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)

def parse(data: bytes):
    upper=data.upper()
    if b'<!DOCTYPE' in upper or b'<!ENTITY' in upper:
        raise BrandError('DTD/entity-bearing XML is unsupported')
    try: return minidom.parseString(data)
    except Exception as exc: raise BrandError(f'Invalid XML: {exc}') from exc

def elems(node, local: str, prefix='a'):
    return list(node.getElementsByTagNameNS(NS[prefix], local))

def children(node, local=None, prefix='a'):
    return [n for n in node.childNodes if n.nodeType==Node.ELEMENT_NODE and
            (local is None or (n.localName==local and n.namespaceURI==NS[prefix]))]

def one(node, local, prefix='a'):
    return next(iter(elems(node,local,prefix)),None)

def direct(node, local, prefix='a'):
    return next(iter(children(node,local,prefix)),None)

def new(doc, prefix, local, attrs=None):
    n=doc.createElementNS(NS[prefix],f'{prefix}:{local}')
    for k,v in (attrs or {}).items():n.setAttribute(k,str(v))
    return n

def text(node) -> str:
    return ''.join(c.data for c in node.childNodes if c.nodeType in (Node.TEXT_NODE,Node.CDATA_SECTION_NODE))

def serial(doc) -> bytes:
    return doc.toxml(encoding='UTF-8')

def hex_colour(value: str) -> str:
    if not isinstance(value,str) or not HEX.fullmatch(value):
        raise BrandError('Colours must be six hexadecimal digits without #')
    return value.upper()

def lum(c):
    vals=[int(c[i:i+2],16)/255 for i in (0,2,4)]
    vals=[v/12.92 if v<=.04045 else ((v+.055)/1.055)**2.4 for v in vals]
    return sum(a*b for a,b in zip(vals,(.2126,.7152,.0722)))

def contrast(a,b):
    x,y=sorted((lum(a),lum(b)))
    return (y+.05)/(x+.05)

class Package:
    def __init__(self, path: Path):
        self.path=Path(path).resolve()
        if self.path.stat().st_size>MAX_TOTAL:raise BrandError('Input file resource limit exceeded')
        self.input_sha256=digest(self.path.read_bytes())
        self.data={};self.infos=[];self.docs={}
        try:
            with zipfile.ZipFile(self.path) as z:
                infos=z.infolist()
                if len(infos)>MAX_ENTRIES:raise BrandError('Too many archive entries')
                total=0
                for i in infos:
                    n=i.filename;p=PurePosixPath(n)
                    if p.is_absolute() or '..' in p.parts or '\\' in n or ':' in n:
                        raise BrandError(f'Unsafe archive path: {n}')
                    if n in self.data:raise BrandError(f'Duplicate archive entry: {n}')
                    if (i.external_attr>>16)&0o170000==0o120000:raise BrandError('Symlink entry unsupported')
                    if i.flag_bits&1:raise BrandError('Encrypted ZIP unsupported')
                    total+=i.file_size
                    if i.file_size>MAX_ENTRY or total>MAX_TOTAL:raise BrandError('Archive resource limit exceeded')
                    self.infos.append(copy.copy(i)); self.data[n]=z.read(i)
        except zipfile.BadZipFile as exc:raise BrandError('Not an unencrypted PPTX ZIP package') from exc
        for n in ['[Content_Types].xml','ppt/presentation.xml']:
            if n not in self.data:raise BrandError(f'Missing PPTX part: {n}')
        # Transitional OOXML only. minidom retains all namespace declarations,
        # including prefixes referred to only by mc:Ignorable attributes.
        if not elems(self.doc('ppt/presentation.xml'),'presentation','p'):
            raise BrandError('Only Transitional OOXML .pptx is supported')

    def assert_writable(self):
        if self.path.suffix.lower()!='.pptx':raise BrandError('Only .pptx inputs can be changed')
        lowered=[p.lower() for p in self.data]
        if any('vbaproject' in p or p.startswith('_xmlsignatures/') for p in lowered):
            raise BrandError('Macro-bearing and signed presentations are read-only')
        if any(p.startswith('ppt/fonts/') for p in lowered):
            raise BrandError('Embedded-font presentations are not repackaged by this tool')
        if b'macroEnabled' in self.data['[Content_Types].xml']:
            raise BrandError('Macro-enabled content types are read-only')

    def doc(self,part):
        if part not in self.data:raise BrandError(f'Missing part: {part}')
        if part not in self.docs:self.docs[part]=parse(self.data[part])
        return self.docs[part]

    def save_doc(self,part):
        doc=self.docs[part]
        if part.startswith('ppt/') and not part.endswith('.rels'):
            for prefix in ('a','p','r'):
                attr='xmlns:'+prefix
                current=doc.documentElement.getAttribute(attr)
                if current and current!=NS[prefix]:raise BrandError('Conflicting namespace prefix: '+prefix)
                if not current:doc.documentElement.setAttribute(attr,NS[prefix])
        self.data[part]=serial(doc)

    def slide_order(self):
        doc=self.doc('ppt/presentation.xml');rels=self.rels('ppt/presentation.xml')
        result=[]
        for s in elems(doc,'sldId','p'):
            rid=s.getAttributeNS(NS['r'],'id'); rel=rels.get(rid)
            if not rel:raise BrandError('Presentation references an absent slide relationship')
            result.append(resolve('ppt/presentation.xml',rel['target']))
        return result

    def rels(self,part):
        rp=rels_path(part)
        if rp not in self.data:return {}
        return {r.getAttribute('Id'):{'target':r.getAttribute('Target'),'type':r.getAttribute('Type'),
                'external':r.getAttribute('TargetMode')=='External'}
                for r in elems(self.doc(rp),'Relationship','rel')}

    def write(self,target: Path,force=False):
        self.assert_writable();target=Path(target).resolve()
        if target==self.path:raise BrandError('In-place editing is prohibited')
        if target.suffix.lower()!='.pptx':raise BrandError('Output must end in .pptx')
        if target.exists() and not force:raise BrandError(f'Refusing to overwrite {target}')
        target.parent.mkdir(parents=True,exist_ok=True)
        fd,tmp=tempfile.mkstemp(prefix='.fortinet-',suffix='.pptx',dir=target.parent);os.close(fd)
        try:
            with zipfile.ZipFile(tmp,'w',compression=zipfile.ZIP_DEFLATED) as z:
                written=set()
                for i in self.infos:
                    z.writestr(i,self.data[i.filename]);written.add(i.filename)
                for n in sorted(set(self.data)-written):z.writestr(n,self.data[n])
            with zipfile.ZipFile(tmp) as z:
                bad=z.testzip()
                if bad:raise BrandError(f'ZIP integrity failure: {bad}')
            os.replace(tmp,target)
        finally:
            if os.path.exists(tmp):os.unlink(tmp)

def rels_path(part):return posixpath.join(posixpath.dirname(part),'_rels',posixpath.basename(part)+'.rels')

def resolve(part,target):
    if target.startswith('/'):return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(part),target))

def identity(shape):
    n=one(shape,'cNvPr','p')
    if n is None:return None
    return {'shape_id':int(n.getAttribute('id')),'name':n.getAttribute('name'),'description':n.getAttribute('descr')}

def shape_nodes(doc):
    return [n for n in doc.getElementsByTagName('*') if n.namespaceURI==NS['p'] and
            n.localName in ('sp','pic','graphicFrame','cxnSp')]

def is_grouped(shape):
    p=shape.parentNode
    while p is not None:
        if p.nodeType==Node.ELEMENT_NODE and p.namespaceURI==NS['p'] and p.localName=='grpSp':return True
        p=p.parentNode
    return False

def box(shape):
    xf=one(shape,'xfrm') or one(shape,'xfrm','p')
    if xf is None:return None
    off=direct(xf,'off');ext=direct(xf,'ext')
    if off is None or ext is None:return None
    return [int(off.getAttribute('x'))/EMU,int(off.getAttribute('y'))/EMU,
            int(ext.getAttribute('cx'))/EMU,int(ext.getAttribute('cy'))/EMU]

def shape_by_id(doc,sid):
    result=[s for s in shape_nodes(doc) if identity(s) and identity(s)['shape_id']==sid]
    if len(result)!=1:raise BrandError(f'Shape ID {sid} is missing or ambiguous')
    return result[0]

def fingerprint(pkg):
    """Visible/stored text, hyperlinks, all opaque binaries. Used before and after."""
    texts={};links={};data_values={}
    for n,data in pkg.data.items():
        if n.endswith('.xml') and n.startswith('ppt/'):
            d=pkg.doc(n);ts=[text(t) for t in elems(d,'t')]
            if ts:texts[n]=ts
            values=[(v.namespaceURI,v.localName,text(v)) for v in d.getElementsByTagName('*')
                    if v.namespaceURI in ('http://schemas.openxmlformats.org/drawingml/2006/chart','http://schemas.openxmlformats.org/officeDocument/2006/math')
                    and v.localName in ('v','f','t')]
            if values:data_values[n]=values
        if n.endswith('.rels'):
            d=pkg.doc(n)
            items=[(r.getAttribute('Id'),r.getAttribute('Target'),r.getAttribute('TargetMode'))
                   for r in elems(d,'Relationship','rel') if r.getAttribute('Type').endswith('/hyperlink')]
            if items:links[n]=items
    return {'text':texts,'hyperlinks':links,'chart_math_values':data_values}

def audit(pkg: Package, expected_font='Inter'):
    tokens=jd(ROOT/'assets/brand-tokens.json');approved=set(tokens['colours'].values())
    size=one(pkg.doc('ppt/presentation.xml'),'sldSz','p')
    w,h=int(size.getAttribute('cx'))/EMU,int(size.getAttribute('cy'))/EMU
    findings=[];slides=[];fonts=set();colours=set();external=[]
    for n in sorted(pkg.data):
        if STYLE_PART.fullmatch(n):
            d=pkg.doc(n)
            for f in elems(d,'latin'):
                v=f.getAttribute('typeface')
                if v:fonts.add(v)
            for c in elems(d,'srgbClr'):
                v=c.getAttribute('val').upper()
                if HEX.fullmatch(v):colours.add(v)
        if n.endswith('.rels'):
            for r in elems(pkg.doc(n),'Relationship','rel'):
                if r.getAttribute('TargetMode')=='External':
                    external.append({'part':n,'id':r.getAttribute('Id'),'type':r.getAttribute('Type'),'target':r.getAttribute('Target')})
    for idx,part in enumerate(pkg.slide_order(),1):
        d=pkg.doc(part);items=[]
        for s in shape_nodes(d):
            ident=identity(s)
            if ident is None:continue
            b=box(s);ts=[text(t) for t in elems(s,'t')]
            item={**ident,'kind':s.localName,'grouped':is_grouped(s),'text':'\n'.join(ts),'box_inches':b}
            sizes=[int(p.getAttribute('sz'))/100 for tag in ['rPr','defRPr','endParaRPr'] for p in elems(s,tag) if p.hasAttribute('sz')]
            if sizes:item['minimum_explicit_font_pt']=min(sizes)
            if sizes and min(sizes)<10:
                findings.append({'code':'SMALL_TEXT','slide':idx,'shape_id':ident['shape_id'],'points':min(sizes),'severity':'review'})
            if b and not item['grouped'] and (b[0]<-0.005 or b[1]<-0.005 or b[0]+b[2]>w+.005 or b[1]+b[3]>h+.005):
                findings.append({'code':'OUT_OF_BOUNDS','slide':idx,'shape_id':ident['shape_id'],'severity':'review'})
            if s.localName=='pic':
                bl=one(s,'blip')
                rid=bl.getAttributeNS(NS['r'],'embed') if bl else ''
                rel=pkg.rels(part).get(rid)
                if rel and not rel['external']:
                    target=resolve(part,rel['target']);item['media_part']=target
                    if target in pkg.data:item['media_sha256']=digest(pkg.data[target])
            items.append(item)
        slides.append({'slide_number':idx,'part':part,'shapes':items})
    unexpected=sorted(f for f in fonts if not f.startswith('+') and f!=expected_font and not any(s in f.lower() for s in SYMBOL_FONTS))
    if unexpected:findings.append({'code':'FONT_DECLARATIONS','fonts':unexpected,'severity':'review'})
    if colours-approved:findings.append({'code':'NON_PALETTE_DIRECT_COLOURS','colours':sorted(colours-approved),'severity':'review','note':'May be user semantics; do not automatically recolour imagery.'})
    return {'schema_version':1,'tool_version':VERSION,'input_sha256':pkg.input_sha256,'slide_size_inches':[w,h],
            'slide_count':len(slides),'declared_latin_fonts':sorted(fonts),'direct_colour_values':sorted(colours),
            'external_relationships':external,'findings':findings,'slides':slides,
            'scope_note':'Structural inventory, not visual or accessibility certification. Inherited fonts, rendered fit, group transforms, theme colour resolution, contrast, and master/layout appearance still need rendering. External relationships are listed, never followed.'}

def set_font(pr,font):
    old=direct(pr,'latin')
    if old is not None and any(s in old.getAttribute('typeface').lower() for s in SYMBOL_FONTS):return 0
    if old is not None:
        if old.getAttribute('typeface')==font:return 0
        old.setAttribute('typeface',font);return 1
    f=new(pr.ownerDocument,'a','latin',{'typeface':font})
    later=[c for c in children(pr) if c.localName in ('ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst')]
    if later:pr.insertBefore(f,later[0])
    else:pr.appendChild(f)
    return 1

def set_solid(pr,colour):
    doc=pr.ownerDocument
    names=('noFill','solidFill','gradFill','blipFill','pattFill','grpFill')
    for n in children(pr):
        if n.namespaceURI==NS['a'] and n.localName in names:pr.removeChild(n)
    fill=new(doc,'a','solidFill');fill.appendChild(new(doc,'a','srgbClr',{'val':hex_colour(colour)}))
    # OOXML child order for character properties or shape/table properties.
    if pr.localName in ('rPr','defRPr','endParaRPr'):
        later=('effectLst','effectDag','highlight','uLnTx','uLn','uFillTx','uFill','latin','ea','cs','sym','hlinkClick','hlinkMouseOver','rtl','extLst')
    else:later=('ln','effectLst','effectDag','scene3d','sp3d','extLst')
    successor=next((c for c in children(pr) if c.localName in later),None)
    if successor:pr.insertBefore(fill,successor)
    else:pr.appendChild(fill)

def run_properties(shape):
    ps=[]
    for r in elems(shape,'r')+elems(shape,'fld'):
        pr=direct(r,'rPr')
        if pr is None:
            pr=new(r.ownerDocument,'a','rPr');r.insertBefore(pr,r.firstChild)
        ps.append(pr)
    ps+=elems(shape,'defRPr')+elems(shape,'endParaRPr')
    return ps

def change_box(shape,values):
    if is_grouped(shape):raise BrandError('Explicit geometry changes inside groups are unsupported')
    if not isinstance(values,list) or len(values)!=4 or not all(isinstance(v,(int,float)) and math.isfinite(v) for v in values):
        raise BrandError('box_inches must contain four finite numbers')
    if values[2]<=0 or values[3]<=0:raise BrandError('Box dimensions must be positive')
    xf=one(shape,'xfrm') or one(shape,'xfrm','p')
    if xf is None:raise BrandError('Shape geometry is inherited; supply a different target')
    off=direct(xf,'off');ext=direct(xf,'ext')
    if off is None or ext is None:raise BrandError('Explicit shape geometry is required')
    for n,a,v in [(off,'x',values[0]),(off,'y',values[1]),(ext,'cx',values[2]),(ext,'cy',values[3])]:n.setAttribute(a,str(round(v*EMU)))

def colours_allowed(values,tokens):
    for v in values:
        if hex_colour(v) not in set(tokens['colours'].values()):raise BrandError(f'Non-palette target colour: {v}')

def shape_style(shape,op,tokens):
    allowed={'slide_part','shape_id','font_size_pt','font_colour','bold','align','fill_colour','line_colour','box_inches'}
    if set(op)-allowed:raise BrandError(f'Unknown shape-style fields: {sorted(set(op)-allowed)}')
    for c in ['font_colour','fill_colour','line_colour']:
        if c in op:colours_allowed([op[c]],tokens)
    if 'font_size_pt' in op:
        v=op['font_size_pt']
        if not isinstance(v,(int,float)) or not 6<=v<=200:raise BrandError('font_size_pt must be between 6 and 200')
    if 'bold' in op and not isinstance(op['bold'],bool):raise BrandError('bold must be true or false')
    for pr in run_properties(shape):
        if 'font_size_pt' in op:pr.setAttribute('sz',str(round(op['font_size_pt']*100)))
        if 'font_colour' in op:set_solid(pr,op['font_colour'])
        if 'bold' in op:pr.setAttribute('b','1' if op['bold'] else '0')
    if 'align' in op:
        aligns={'left':'l','center':'ctr','right':'r'}
        if op['align'] not in aligns:raise BrandError('Unsupported alignment')
        for p in elems(shape,'p'):
            pr=direct(p,'pPr')
            if pr is None:pr=new(p.ownerDocument,'a','pPr');p.insertBefore(pr,p.firstChild)
            pr.setAttribute('algn',aligns[op['align']])
    sp=direct(shape,'spPr','p')
    if 'fill_colour' in op:
        if sp is None:raise BrandError('No explicit shape properties for fill')
        set_solid(sp,op['fill_colour'])
    if 'line_colour' in op:
        if sp is None:raise BrandError('No explicit shape properties for line')
        ln=direct(sp,'ln')
        if ln is None:
            ln=new(sp.ownerDocument,'a','ln');later=next((c for c in children(sp) if c.localName in ('effectLst','effectDag','scene3d','sp3d','extLst')),None)
            if later:sp.insertBefore(ln,later)
            else:sp.appendChild(ln)
        set_solid(ln,op['line_colour'])
    if 'box_inches' in op:change_box(shape,op['box_inches'])

def style_table(shape,op,tokens):
    allowed={'slide_part','shape_id','header_fill','header_text','body_fill','body_text'}
    if set(op)-allowed:raise BrandError('Unknown table-style fields')
    colours_allowed([op[k] for k in op if k not in ('slide_part','shape_id')],tokens)
    table=one(shape,'tbl')
    if table is None:raise BrandError('Target is not a DrawingML table')
    for i,row in enumerate(children(table,'tr')):
        colour=op.get('header_fill' if i==0 else 'body_fill')
        txt=op.get('header_text' if i==0 else 'body_text')
        for cell in children(row,'tc'):
            if colour:
                pr=direct(cell,'tcPr')
                if pr is None:pr=new(cell.ownerDocument,'a','tcPr');cell.appendChild(pr)
                set_solid(pr,colour)
            if txt:
                for pr in run_properties(cell):set_solid(pr,txt)

# Image replacement is intentionally explicit. Never infer product identity from pixels.
def catalog():return jd(ROOT/'assets/asset-catalog.json')

def asset_for(id):
    candidates=[a for a in catalog()['assets'] if a['id']==id]
    if len(candidates)!=1 or not candidates[0].get('approved'):raise BrandError(f'Unknown/unapproved asset: {id}')
    a=candidates[0];p=(ROOT/a['path']).resolve()
    if not p.is_relative_to(ROOT.resolve()) or not p.is_file():raise BrandError('Asset path escaped package or is missing')
    b=p.read_bytes()
    if digest(b)!=a['sha256']:raise BrandError(f'Asset checksum mismatch: {id}')
    if p.suffix.lower()!='.png' or b[:8]!=b'\x89PNG\r\n\x1a\n':raise BrandError('Replacement assets require a catalogued PNG fallback')
    return a,b

def contain_box(a,target):
    x,y,w,h=target;ratio=a['width_px']/a['height_px']
    if w/h>ratio:
        nw=h*ratio;return [x+(w-nw)/2,y,nw,h]
    nh=w/ratio;return [x,y+(h-nh)/2,w,nh]

def validate_mark(a,b,op):
    if a['kind'] not in ('logo','grid'):return
    if op.get('clear_space_confirmed') is not True:
        raise BrandError('Logo/Grid placement requires explicit clear_space_confirmed=true after reviewing slide, layout, and master')
    if a['kind']=='grid' and op.get('brand_context_confirmed') is not True:
        raise BrandError('Grid-only usage requires brand_context_confirmed=true')
    bg=hex_colour(op.get('background_hex',''))
    ink=a['ink_bbox_px'];sx=b[2]/a['width_px'];sy=b[3]/a['height_px']
    min_width=.75 if a['kind']=='logo' else .25
    if (ink[2]-ink[0])*sx<min_width-1e-6:raise BrandError('Visible logo/Grid is below the print-safe minimum width')
    marks=['DA291C'] if a['id'].endswith('red') else []
    if 'black' in a['id']:marks+=['000000']
    if 'white' in a['id']:marks+=['FFFFFF']
    if any(contrast(c,bg)<3 for c in marks):raise BrandError('Logo contrast below 3:1 for declared background')
    # The required clear zone is reported for human review. The existing transparent
    # padding is not assumed to equal the entire cap-height clearance.
    height=(ink[3]-ink[1])*sy;clear=height*(1 if a['kind']=='logo' else .5)
    return [b[0]+ink[0]*sx-clear,b[1]+ink[1]*sy-clear,(ink[2]-ink[0])*sx+2*clear,height+2*clear]

def add_image_relationship(pkg,part,asset_bytes):
    target=f'ppt/media/fortinet-{digest(asset_bytes)[:24]}.png'
    pkg.data[target]=asset_bytes
    rp=rels_path(part)
    if rp not in pkg.data:
        pkg.data[rp]=f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{NS["rel"]}"/>'.encode()
    doc=pkg.doc(rp);ids={r.getAttribute('Id') for r in elems(doc,'Relationship','rel')}
    i=1
    while f'rIdFortinet{i}' in ids:i+=1
    rid=f'rIdFortinet{i}'
    r=doc.createElementNS(NS['rel'],'Relationship')
    for k,v in {'Id':rid,'Type':NS['r']+'/image','Target':posixpath.relpath(target,posixpath.dirname(part))}.items():r.setAttribute(k,v)
    doc.documentElement.appendChild(r);pkg.save_doc(rp)
    ct=pkg.doc('[Content_Types].xml')
    defaults=elems(ct,'Default','ct')
    if not any(d.getAttribute('Extension').lower()=='png' for d in defaults):
        el=ct.createElementNS(NS['ct'],'Default');el.setAttribute('Extension','png');el.setAttribute('ContentType','image/png');ct.documentElement.appendChild(el);pkg.save_doc('[Content_Types].xml')
    return rid

def image_op(pkg,op,adding=False):
    allowed={'slide_part','shape_id','asset_id','box_inches','background_hex','clear_space_confirmed','brand_context_confirmed'}
    if set(op)-allowed:raise BrandError('Unknown image-operation fields')
    part=op['slide_part'];doc=pkg.doc(part);a,data=asset_for(op['asset_id'])
    if adding:
        if a['kind'] not in ('logo','grid'):raise BrandError('add_logos only accepts logo/Grid assets')
        target=op.get('box_inches')
        if target is None:raise BrandError('Logo insertion requires an explicit reserved box')
        ids=[identity(s)['shape_id'] for s in shape_nodes(doc) if identity(s)]
        sid=max(ids+[1])+1
        pic=new(doc,'p','pic');nv=new(doc,'p','nvPicPr')
        nv.appendChild(new(doc,'p','cNvPr',{'id':sid,'name':f'Fortinet:{a["id"]}','descr':a['label']}))
        np=new(doc,'p','cNvPicPr');np.appendChild(new(doc,'a','picLocks',{'noChangeAspect':1}));nv.appendChild(np);nv.appendChild(new(doc,'p','nvPr'));pic.appendChild(nv)
        sp=new(doc,'p','spPr');xf=new(doc,'a','xfrm');xf.appendChild(new(doc,'a','off',{'x':0,'y':0}));xf.appendChild(new(doc,'a','ext',{'cx':1,'cy':1}));sp.appendChild(xf)
        geo=new(doc,'a','prstGeom',{'prst':'rect'});geo.appendChild(new(doc,'a','avLst'));sp.appendChild(geo);pic.appendChild(sp)
    else:
        pic=shape_by_id(doc,op['shape_id'])
        if pic.localName!='pic':raise BrandError('replace_images targets an existing picture only, never text or native shapes')
        if is_grouped(pic):raise BrandError('Grouped picture replacement requires manual review; not supported')
        xf=one(pic,'xfrm')
        if xf and any(xf.hasAttribute(v) and xf.getAttribute(v) not in ('0','false') for v in ('rot','flipH','flipV')):
            raise BrandError('Rotated/flipped picture replacement unsupported')
        target=op.get('box_inches') or box(pic)
    if not isinstance(target,list) or len(target)!=4 or not all(isinstance(v,(int,float)) and math.isfinite(v) for v in target) or target[2]<=0 or target[3]<=0:
        raise BrandError('Image target needs four finite box_inches with positive size')
    b=contain_box(a,target);zone=validate_mark(a,b,op)
    size=one(pkg.doc('ppt/presentation.xml'),'sldSz','p');sw=int(size.getAttribute('cx'))/EMU;sh=int(size.getAttribute('cy'))/EMU
    z=zone or b
    if z[0]<0 or z[1]<0 or z[0]+z[2]>sw or z[1]+z[3]>sh:raise BrandError('Asset or required clear-space zone exceeds slide canvas')
    rid=add_image_relationship(pkg,part,data)
    fill=direct(pic,'blipFill','p')
    if fill is not None:pic.removeChild(fill)
    fill=new(doc,'p','blipFill');blip=new(doc,'a','blip');blip.setAttributeNS(NS['r'],'r:embed',rid);fill.appendChild(blip)
    stretch=new(doc,'a','stretch');stretch.appendChild(new(doc,'a','fillRect'));fill.appendChild(stretch)
    pic.insertBefore(fill,direct(pic,'spPr','p'))
    change_box(pic,b)
    pr=one(pic,'cNvPr','p');pr.setAttribute('descr',a['label']+'; source: '+a['provenance']['source'])
    # Clear decorative picture effects; keep unrelated original properties intact.
    sp=direct(pic,'spPr','p')
    for child in list(children(sp)):
        if child.localName in ('effectLst','effectDag','scene3d','sp3d','ln','prstGeom','custGeom'):sp.removeChild(child)
    # Do not clip the official replacement to an old oval/custom picture mask.
    geo=new(doc,'a','prstGeom',{'prst':'rect'});geo.appendChild(new(doc,'a','avLst'))
    following=next((c for c in children(sp) if c.localName!='xfrm'),None)
    if following:sp.insertBefore(geo,following)
    else:sp.appendChild(geo)
    inherited_style=direct(pic,'style','p')
    if inherited_style is not None:pic.removeChild(inherited_style)
    if adding:
        tree=one(doc,'spTree','p');ext=direct(tree,'extLst','p')
        if ext:tree.insertBefore(pic,ext)
        else:tree.appendChild(pic)
    pkg.save_doc(part)
    return {'part':part,'shape_id':int(pr.getAttribute('id')),'asset_id':a['id'],'clear_zone_inches':zone,'human_clear_space_confirmation':op.get('clear_space_confirmed',False)}

def new_plan(pkg,font,theme_colours,mappings):
    colours={}
    for v in mappings:
        if v.count('=')!=1:raise BrandError('Use --map-colour OLDHEX=NEWHEX')
        old,newc=v.split('=');colours[hex_colour(old)]=hex_colour(newc)
    tokens=jd(ROOT/'assets/brand-tokens.json');colours_allowed(colours.values(),tokens)
    return {'schema_version':1,'tool_version':VERSION,'input_sha256':pkg.input_sha256,
            'tokens_sha256':digest((ROOT/'assets/brand-tokens.json').read_bytes()),
            'catalog_sha256':digest((ROOT/'assets/asset-catalog.json').read_bytes()),
            'font':font,'theme_colours':theme_colours,'colour_map':colours,
            'style_shapes':[],'table_styles':[],'replace_images':[],'add_logos':[],
            'review_notes':['Review this plan before apply. Font changes can reflow text. Theme changes may affect charts and inherited colours.',
                            'No automatic icon/product matching, logo insertion, copy rewriting, layout rearrangement, or image recolouring.',
                            'Explicit picture operations require verified catalog IDs. Check clear space against slide, master, and layout.']}

def apply(pkg,plan,output,force=False):
    pkg.assert_writable()
    allowed={'schema_version','tool_version','input_sha256','tokens_sha256','catalog_sha256','font','theme_colours','colour_map','style_shapes','table_styles','replace_images','add_logos','review_notes'}
    if set(plan)-allowed:raise BrandError(f'Unknown plan fields: {sorted(set(plan)-allowed)}')
    if plan.get('schema_version')!=1:raise BrandError('Unsupported plan schema')
    if plan.get('input_sha256')!=pkg.input_sha256:raise BrandError('Input changed after plan creation')
    for key,path in [('tokens_sha256',ROOT/'assets/brand-tokens.json'),('catalog_sha256',ROOT/'assets/asset-catalog.json')]:
        if plan.get(key)!=digest(path.read_bytes()):raise BrandError(f'{key} changed after plan creation')
    font=plan.get('font')
    if font not in ('Inter','Arial'):raise BrandError('Font profile must be Inter or explicit legacy Arial')
    if not isinstance(plan.get('theme_colours'),bool):raise BrandError('theme_colours must be boolean')
    tokens=jd(ROOT/'assets/brand-tokens.json');cmap=plan.get('colour_map',{})
    if not isinstance(cmap,dict):raise BrandError('colour_map must be an object')
    for k,v in cmap.items():hex_colour(k);colours_allowed([v],tokens)
    cmap={k.upper():v.upper() for k,v in cmap.items()}
    for field in ('style_shapes','table_styles','replace_images','add_logos'):
        if not isinstance(plan.get(field),list):raise BrandError(f'{field} must be an array')
    original=dict(pkg.data);before=fingerprint(pkg);changes=[];image_report=[]
    for part in sorted(pkg.data):
        if not STYLE_PART.fullmatch(part):continue
        d=pkg.doc(part);prior=serial(d);count=0
        for local in ('rPr','defRPr','endParaRPr'):
            for pr in elems(d,local):count+=set_font(pr,font)
        # Theme major/minor Latin declarations and any remaining direct Latin refs.
        for f in elems(d,'latin'):
            if any(s in f.getAttribute('typeface').lower() for s in SYMBOL_FONTS):continue
            if f.getAttribute('typeface')!=font:f.setAttribute('typeface',font);count+=1
        for c in elems(d,'srgbClr'):
            v=c.getAttribute('val').upper()
            if v in cmap:c.setAttribute('val',cmap[v]);count+=1
        if plan['theme_colours'] and part.startswith('ppt/theme/'):
            for scheme in elems(d,'clrScheme'):
                for slot in children(scheme):
                    value=tokens['theme'].get(slot.localName)
                    if value:
                        for child in list(slot.childNodes):slot.removeChild(child)
                        slot.appendChild(new(d,'a','srgbClr',{'val':value}))
        if serial(d)!=prior:pkg.save_doc(part);changes.append({'part':part,'font_or_colour_edits':count})
    slides=set(pkg.slide_order())
    for field,fn in [('style_shapes',shape_style),('table_styles',style_table)]:
        for op in plan[field]:
            if op.get('slide_part') not in slides:raise BrandError('Operations must address an actual slide part from audit')
            if not isinstance(op.get('shape_id'),int):raise BrandError('shape_id must be an integer')
            part=op['slide_part'];s=shape_by_id(pkg.doc(part),op['shape_id']);fn(s,op,tokens);pkg.save_doc(part)
    for field,adding in [('replace_images',False),('add_logos',True)]:
        for op in plan[field]:
            if op.get('slide_part') not in slides:raise BrandError('Image operation is outside actual slides')
            image_report.append(image_op(pkg,op,adding))
    after=fingerprint(pkg)
    if before!=after:raise BrandError('Invariant violation: stored text or hyperlink targets changed')
    # Validate local relationship targets without following external URLs.
    broken=[]
    for part in list(pkg.data):
        if not part.endswith('.rels'):continue
        owner=posixpath.join(posixpath.dirname(posixpath.dirname(part)),posixpath.basename(part)[:-5])
        if part=='_rels/.rels':owner=''
        for rel in pkg.rels(owner).values():
            if not rel['external'] and resolve(owner,rel['target']) not in pkg.data:broken.append((part,rel['target']))
    if broken:raise BrandError(f'Unresolved local relationships: {broken[:5]}')
    modified=[n for n in original if pkg.data.get(n)!=original[n]]
    added=sorted(set(pkg.data)-set(original))
    # No source entry is removed and binary data remains untouched (new assets only).
    for n,v in original.items():
        if not n.endswith(('.xml','.rels')) and pkg.data[n]!=v:raise BrandError(f'Opaque part changed: {n}')
    pkg.write(output,force)
    result=audit(Package(output),font)
    result['application']={'plan_sha256':digest(json.dumps(plan,sort_keys=True).encode()),'source_sha256':pkg.input_sha256,
            'modified_parts':sorted(modified),'added_parts':added,'all_existing_opaque_parts_byte_identical':True,
            'stored_text_and_hyperlinks_preserved':True,'image_operations':image_report,
            'render_required':True,'publication_ready':False,'warnings':['Structural success does not prove text fit, colour contrast, template compliance, or logo clear-space compliance. Render and inspect all slides.']}
    return result

def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--version',action='version',version=VERSION)
    sub=parser.add_subparsers(dest='command',required=True)
    pa=sub.add_parser('audit',help='Read-only inventory of shapes, fonts, colours, media, and review findings')
    pa.add_argument('input',type=Path);pa.add_argument('--output',type=Path);pa.add_argument('--font',choices=['Inter','Arial'],default='Inter');pa.add_argument('--force',action='store_true')
    pp=sub.add_parser('plan',help='Prepare editable, hash-bound JSON; no presentation changes')
    pp.add_argument('input',type=Path);pp.add_argument('--output',type=Path,required=True);pp.add_argument('--font',choices=['Inter','Arial'],default='Inter')
    pp.add_argument('--theme-colours',action='store_true');pp.add_argument('--map-colour',action='append',default=[]);pp.add_argument('--force',action='store_true')
    ap=sub.add_parser('apply',help='Apply a reviewed plan to a new .pptx, never the source')
    ap.add_argument('input',type=Path);ap.add_argument('--plan',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);ap.add_argument('--report',type=Path,required=True);ap.add_argument('--force',action='store_true')
    ip=sub.add_parser('icons',help='Query the local verified asset catalog without opening artwork')
    ip.add_argument('--query',default='');ip.add_argument('--output',type=Path);ip.add_argument('--force',action='store_true')
    args=parser.parse_args(argv)
    try:
        if args.command=='icons':
            q=args.query.casefold();rows=[a for a in catalog()['assets'] if q in json.dumps([a['id'],a['label'],a['tags']]).casefold()]
            result={'count':len(rows),'assets':rows}
        else:
            pkg=Package(args.input)
            if args.command=='audit':result=audit(pkg,args.font)
            elif args.command=='plan':result=new_plan(pkg,args.font,args.theme_colours,args.map_colour)
            else:
                if args.report.resolve() in (args.input.resolve(),args.output.resolve(),args.plan.resolve()):raise BrandError('Report path must be separate from input, output, and plan')
                if args.report.exists() and not args.force:raise BrandError('Report exists; choose a new report path')
                result=apply(pkg,jd(args.plan),args.output,args.force)
                write_json(args.report,result,args.force)
                print(json.dumps({'output':str(args.output),'report':str(args.report),'text_preserved':True,'render_required':True}))
                return 0
        if args.output:
            if args.command!='icons' and args.output.resolve()==args.input.resolve():raise BrandError('Output must not overwrite source')
            write_json(args.output,result,args.force)
            print(json.dumps({'output':str(args.output),'command':args.command}))
        else:print(json.dumps(result,indent=2,ensure_ascii=False))
        return 0
    except (BrandError,OSError,ValueError,KeyError,TypeError) as exc:
        print(f'ERROR: {exc}',file=sys.stderr);return 2

if __name__=='__main__':sys.exit(main())
