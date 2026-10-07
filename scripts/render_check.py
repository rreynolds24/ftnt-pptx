#!/usr/bin/env python3
"""Optional local PPTX rendering and PDF text/font checks. No network calls.
Requires a locally installed LibreOffice plus PyMuPDF. Does not install software.
All slides are rendered to PNG; inspect them manually before release.
"""
from __future__ import annotations
import argparse,json,os,shutil,subprocess,sys,tempfile
from pathlib import Path
from fortinet_pptx import Package,BrandError,audit

def run(input_path,outdir,font='Inter',width=2560,timeout=180,expected_terms=None):
    import fitz
    pkg=Package(input_path);pkg.assert_writable()
    # Hyperlinks are not fetched by this tool; linked media/data could be refreshed
    # by a renderer, so refuse that dependency before invoking LibreOffice.
    for rel in audit(pkg,font)['external_relationships']:
        if not rel['type'].endswith('/hyperlink'):
            raise BrandError('Offline rendering refuses externally linked media or data: '+rel['type'])
    if outdir.exists():raise BrandError('Render output must be a new directory')
    exe=shutil.which('soffice') or shutil.which('libreoffice')
    if not exe:
        for c in [Path(os.environ.get('PROGRAMFILES','C:/Program Files'))/'LibreOffice/program/soffice.exe',Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')]:
            if c.is_file():exe=str(c);break
    if not exe:raise BrandError('LibreOffice is not installed or on PATH')
    if not 640<=width<=7680:raise BrandError('PNG width must be between 640 and 7680')
    outdir.mkdir(parents=True)
    with tempfile.TemporaryDirectory(prefix='fortinet-lo-') as tmp:
        profile=Path(tmp)/'profile';profile.mkdir()
        args=[exe,'-env:UserInstallation='+profile.as_uri(),'--headless','--norestore','--convert-to','pdf','--outdir',str(outdir.resolve()),str(input_path.resolve())]
        proc=subprocess.run(args,capture_output=True,text=True,timeout=timeout,check=False)
        (outdir/'render.log').write_text(proc.stdout+'\n'+proc.stderr)
    pdf=outdir/(input_path.stem+'.pdf')
    if proc.returncode or not pdf.exists() or pdf.stat().st_size==0:raise BrandError('LibreOffice did not produce a nonempty PDF; inspect render.log')
    d=fitz.open(pdf);fontnames=set();texts=[];pages=[]
    for n,page in enumerate(d,1):
        page.get_pixmap(matrix=fitz.Matrix(width/page.rect.width,width/page.rect.width),alpha=False).save(outdir/f'slide-{n}.png')
        ft={v[3] for v in page.get_fonts(full=True)};fontnames|=ft
        texts.append(page.get_text());pages.append({'page':n,'size_pt':list(page.rect),'font_names':sorted(ft),'text_characters':len(texts[-1])})
    names=[f.split('+')[-1] for f in fontnames]
    unexpected=[f for f in names if font.casefold() not in f.casefold() and not any(t in f.casefold() for t in ('symbol','wingdings','webdings'))]
    combined='\n'.join(texts)
    missing=[t for t in (expected_terms or []) if t not in combined]
    count_ok=len(d)==len(pkg.slide_order())
    result={'tool_version':'2.0.0','rendered_pages':len(d),'slide_count_matches':count_ok,'font_names':sorted(fontnames),'unexpected_rendered_fonts':unexpected,'missing_required_terms':missing,'pages':pages,'structural_audit':audit(pkg,font),'automated_checks_pass':count_ok and not unexpected and not missing,'visual_review_required':True,'manual_visual_review_complete':False,'note':'Font-name and text extraction checks are evidence of this renderer output, not proof of installed font licensing, overall layout quality, or PowerPoint/Keynote parity.'}
    (outdir/'render-validation.json').write_text(json.dumps(result,indent=2)+'\n')
    return result

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('input',type=Path);p.add_argument('--output-dir',type=Path,required=True);p.add_argument('--font',default='Inter');p.add_argument('--width',type=int,default=2560);p.add_argument('--timeout',type=int,default=180);p.add_argument('--expected-terms',type=Path)
    a=p.parse_args()
    try:
        r=run(a.input,a.output_dir,a.font,a.width,a.timeout,json.loads(a.expected_terms.read_text()) if a.expected_terms else None)
        print(json.dumps({'output':str(a.output_dir),'automated_checks_pass':r['automated_checks_pass'],'visual_review_required':True}));return 0 if r['automated_checks_pass'] else 3
    except Exception as exc:print('ERROR: '+str(exc),file=sys.stderr);return 2
if __name__=='__main__':sys.exit(main())
