"""Sanitized synthetic OOXML fixtures. These tests do not prove Office renderer parity."""
from __future__ import annotations
import base64,copy,hashlib,json,subprocess,sys,tempfile,unittest,zipfile
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts'))
import fortinet_pptx as f
import brand_assets as ba
P=f.NS['p'];A=f.NS['a'];R=f.NS['r'];REL=f.NS['rel'];CT=f.NS['ct']
PNG=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2lKQAAAAASUVORK5CYII=')

def picture(id=3,rid='rId2'):
 return f'''<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="Picture {id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="{rid}"/><a:srcRect l="1000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="100000" y="3000000"/><a:ext cx="1600000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'''

def fixture(path,changes=None,extras=None):
 files={
 '[Content_Types].xml':f'''<Types xmlns="{CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>''',
 '_rels/.rels':f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
 'ppt/presentation.xml':f'''<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:defaultTextStyle><a:defPPr><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></a:defPPr></p:defaultTextStyle></p:presentation>''',
 'ppt/_rels/presentation.xml.rels':f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/slide" Target="slides/slide1.xml"/></Relationships>',
 'ppt/slides/slide1.xml':f'''<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" mc:Ignorable="p14"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="1000000"/><a:ext cx="6000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EE3124"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-AU" sz="2400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/><a:hlinkClick r:id="rId1"/></a:rPr><a:t>Preserve this factual text: 100 Gbps</a:t></a:r><a:r><a:rPr><a:latin typeface="Wingdings"/></a:rPr><a:t>X</a:t></a:r><a:endParaRPr lang="en-AU"/></a:p></p:txBody></p:sp>{picture()}{picture(4)}<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Test Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="4000000" y="3000000"/><a:ext cx="4000000" cy="1200000"/></p:xfrm><a:graphic><a:graphicData uri="{A}/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="4000000"/></a:tblGrid><a:tr h="600000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Header</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="600000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr/><a:t>Row: unchanged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>''',
 'ppt/slides/_rels/slide1.xml.rels':f'<Relationships xmlns="{REL}"><Relationship Id="rId1" Type="{R}/hyperlink" Target="https://example.invalid/keep?mode=1" TargetMode="External"/><Relationship Id="rId2" Type="{R}/image" Target="../media/shared.png"/></Relationships>',
 'ppt/theme/theme1.xml':f'<a:theme xmlns:a="{A}" name="Fixture"><a:themeElements><a:clrScheme name="Old"><a:accent1><a:srgbClr val="EE3124"/></a:accent1></a:clrScheme><a:fontScheme name="Old"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface="EastAsianOriginal"/><a:cs typeface="ComplexScriptOriginal"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>',
 'ppt/notesSlides/notesSlide1.xml':f'<p:notes xmlns:p="{P}" xmlns:a="{A}"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:rPr/><a:t>Original notes; not rewritten.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>',
 'ppt/media/shared.png':PNG,
 'ppt/embeddings/fixture.bin':b'OPAQUE-DO-NOT-CHANGE',
 'customXml/item1.xml':'<custom preserve="true">External metadata remains intact</custom>',
 }
 files.update(changes or {});files.update(extras or {})
 with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
  for n,b in files.items():z.writestr(n,b.encode() if isinstance(b,str) else b)
 return path

class OfflineTests(unittest.TestCase):
 def setUp(self):
  self.tmp=tempfile.TemporaryDirectory();self.dir=Path(self.tmp.name);self.src=fixture(self.dir/'source.pptx');self.pkg=f.Package(self.src);self.plan=f.new_plan(self.pkg,'Inter',False,[]);self.out=self.dir/'output.pptx'
 def tearDown(self):self.tmp.cleanup()
 def go(self):return f.apply(self.pkg,self.plan,self.out)
 def test_audit_no_write(self):
  h=f.digest(self.src.read_bytes());r=f.audit(self.pkg);self.assertEqual(r['slide_count'],1);self.assertEqual(h,f.digest(self.src.read_bytes()));self.assertFalse(self.out.exists())
 def test_plan_defaults_non_destructive(self):
  self.assertFalse(self.plan['theme_colours']);self.assertEqual(self.plan['replace_images'],[]);self.assertEqual(self.plan['add_logos'],[])
 def test_text_links_and_notes_preserved(self):
  before=f.fingerprint(self.pkg);r=self.go();self.assertEqual(before,f.fingerprint(f.Package(self.out)));self.assertTrue(r['application']['stored_text_and_hyperlinks_preserved'])
 def test_source_immutable(self):
  h=f.digest(self.src.read_bytes());self.go();self.assertEqual(h,f.digest(self.src.read_bytes()))
 def test_no_in_place_even_force(self):
  with self.assertRaises(f.BrandError):f.apply(self.pkg,self.plan,self.src,True)
 def test_existing_output_refused(self):
  self.out.write_bytes(b'old');
  with self.assertRaises(f.BrandError):self.go()
  self.assertEqual(self.out.read_bytes(),b'old')
 def test_input_plan_hash(self):
  self.plan['input_sha256']='0'*64
  with self.assertRaises(f.BrandError):self.go()
  self.assertFalse(self.out.exists())
 def test_token_hash(self):
  self.plan['tokens_sha256']='0'*64
  with self.assertRaises(f.BrandError):self.go()
 def test_catalog_hash(self):
  self.plan['catalog_sha256']='0'*64
  with self.assertRaises(f.BrandError):self.go()
 def test_unknown_fields_fail(self):
  self.plan['rewrite_content']=True
  with self.assertRaises(f.BrandError):self.go()
 def test_font_declared_and_symbols_preserved(self):
  self.go();d=f.Package(self.out).doc('ppt/slides/slide1.xml');fonts=[x.getAttribute('typeface') for x in f.elems(d,'latin')];self.assertIn('Inter',fonts);self.assertIn('Wingdings',fonts);self.assertNotIn('Calibri',fonts)
 def test_non_latin_font_declarations_preserved(self):
  self.go();d=f.Package(self.out).doc('ppt/theme/theme1.xml');self.assertEqual(f.one(d,'ea').getAttribute('typeface'),'EastAsianOriginal')
 def test_presentation_default_font(self):
  self.go();d=f.Package(self.out).doc('ppt/presentation.xml');self.assertEqual(f.one(d,'latin').getAttribute('typeface'),'Inter')
 def test_namespace_prefix_only_used_in_ignorable_preserved(self):
  self.go();d=f.Package(self.out).doc('ppt/slides/slide1.xml');self.assertTrue(d.documentElement.hasAttribute('xmlns:p14'));self.assertEqual(d.documentElement.getAttribute('mc:Ignorable'),'p14')
 def test_opaque_data_preserved(self):
  old=self.pkg.data['ppt/embeddings/fixture.bin'];self.go();self.assertEqual(f.Package(self.out).data['ppt/embeddings/fixture.bin'],old)
 def test_unknown_xml_untouched(self):
  old=self.pkg.data['customXml/item1.xml'];self.go();self.assertEqual(f.Package(self.out).data['customXml/item1.xml'],old)
 def test_theme_recolour_opt_in(self):
  self.plan['theme_colours']=True;self.go();self.assertEqual(f.one(f.Package(self.out).doc('ppt/theme/theme1.xml'),'srgbClr').getAttribute('val'),'DA291C')
 def test_explicit_direct_colour_map(self):
  self.plan['colour_map']={'EE3124':'DA291C'};self.go();d=f.Package(self.out).doc('ppt/slides/slide1.xml');self.assertNotIn('EE3124',[c.getAttribute('val') for c in f.elems(d,'srgbClr')])
 def test_non_brand_target_rejected(self):
  self.plan['colour_map']={'EE3124':'FF00FF'}
  with self.assertRaises(f.BrandError):self.go()
 def test_shape_style_explicit(self):
  self.plan['style_shapes']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':2,'font_size_pt':22,'bold':True,'font_colour':'000000','fill_colour':'F0F0F0','line_colour':'DA291C','align':'left'}];self.go();d=f.Package(self.out).doc('ppt/slides/slide1.xml');self.assertEqual(f.one(d,'rPr').getAttribute('sz'),'2200')
 def test_geometry_explicit(self):
  self.plan['style_shapes']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':2,'box_inches':[1,1,6,1]}];self.go();self.assertEqual(f.box(f.shape_by_id(f.Package(self.out).doc('ppt/slides/slide1.xml'),2)),[1,1,6,1])
 def test_invalid_shape_target(self):
  self.plan['style_shapes']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':999,'bold':True}]
  with self.assertRaises(f.BrandError):self.go()
 def test_table_header_style(self):
  self.plan['table_styles']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':5,'header_fill':'DA291C','header_text':'FFFFFF','body_fill':'FFFFFF','body_text':'000000'}];self.go();d=f.Package(self.out).doc('ppt/slides/slide1.xml');tbl=f.one(d,'tbl');self.assertIn('DA291C',[c.getAttribute('val') for c in f.elems(tbl,'srgbClr')])
 def test_only_selected_image_instance_changes(self):
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'secure-networking'}];self.go();p=f.Package(self.out);d=p.doc('ppt/slides/slide1.xml');a=f.one(f.shape_by_id(d,3),'blip').getAttributeNS(R,'embed');b=f.one(f.shape_by_id(d,4),'blip').getAttributeNS(R,'embed');self.assertNotEqual(a,b);self.assertEqual(b,'rId2');self.assertEqual(p.data['ppt/media/shared.png'],PNG)
 def test_image_crop_removed(self):
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'secure-networking'}];self.go();self.assertFalse(f.elems(f.shape_by_id(f.Package(self.out).doc('ppt/slides/slide1.xml'),3),'srcRect'))
 def test_picture_operation_never_deletes_text(self):
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':2,'asset_id':'secure-networking'}]
  with self.assertRaises(f.BrandError):self.go()
 def test_unknown_asset_rejected(self):
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'invented-product'}]
  with self.assertRaises(f.BrandError):self.go()
 def test_logo_clearance_requires_confirmation(self):
  self.plan['add_logos']=[{'slide_part':'ppt/slides/slide1.xml','asset_id':'fortinet-logo-rgb-black-red','box_inches':[10,.1,2.5,.75],'background_hex':'FFFFFF'}]
  with self.assertRaises(f.BrandError):self.go()
 def test_logo_insertion_preserves_content(self):
  self.plan['add_logos']=[{'slide_part':'ppt/slides/slide1.xml','asset_id':'fortinet-logo-rgb-black-red','box_inches':[10,.1,2.5,.75],'background_hex':'FFFFFF','clear_space_confirmed':True}];r=self.go();self.assertEqual(len(r['application']['image_operations']),1)
 def test_logo_low_contrast_rejected(self):
  self.plan['add_logos']=[{'slide_part':'ppt/slides/slide1.xml','asset_id':'fortinet-logo-rgb-black-red','box_inches':[10,.1,2.5,.75],'background_hex':'DA291C','clear_space_confirmed':True}]
  with self.assertRaises(f.BrandError):self.go()
 def test_visible_logo_minimum(self):
  self.plan['add_logos']=[{'slide_part':'ppt/slides/slide1.xml','asset_id':'fortinet-logo-rgb-black-red','box_inches':[10,.1,.3,.1],'background_hex':'FFFFFF','clear_space_confirmed':True}]
  with self.assertRaises(f.BrandError):self.go()
 def test_grid_requires_established_brand(self):
  self.plan['add_logos']=[{'slide_part':'ppt/slides/slide1.xml','asset_id':'fortinet-logomark-rgb-red','box_inches':[10,.1,1,1],'background_hex':'FFFFFF','clear_space_confirmed':True}]
  with self.assertRaises(f.BrandError):self.go()
 def test_idempotent_semantics(self):
  self.go();p=f.Package(self.out);p2=f.new_plan(p,'Inter',False,[]);out2=self.dir/'twice.pptx';r=f.apply(p,p2,out2);self.assertEqual(r['application']['modified_parts'],[])
 def test_macro_refused(self):
  fixture(self.src,extras={'ppt/vbaProject.bin':b'test'});p=f.Package(self.src)
  with self.assertRaises(f.BrandError):f.apply(p,f.new_plan(p,'Inter',False,[]),self.out)
 def test_signed_refused(self):
  fixture(self.src,extras={'_xmlsignatures/sig1.xml':b'<test/>'});p=f.Package(self.src)
  with self.assertRaises(f.BrandError):p.assert_writable()
 def test_embedded_font_refused(self):
  fixture(self.src,extras={'ppt/fonts/font1.fntdata':b'test'});p=f.Package(self.src)
  with self.assertRaises(f.BrandError):p.assert_writable()
 def test_zip_traversal_rejected(self):
  fixture(self.src,extras={'../bad':b'x'})
  with self.assertRaises(f.BrandError):f.Package(self.src)
 def test_xml_entities_rejected(self):
  with self.assertRaises(f.BrandError):f.parse(b'<!DOCTYPE x [<!ENTITY e SYSTEM "file:///x">]><x>&e;</x>')
 def test_relationship_missing_target_rejected(self):
  self.pkg.data.pop('ppt/media/shared.png')
  with self.assertRaises(f.BrandError):self.go()
 def test_readonly_external_hyperlink(self):
  r=f.audit(self.pkg);self.assertEqual(r['external_relationships'][0]['target'],'https://example.invalid/keep?mode=1')
 def test_structural_audit_never_claims_release_ready(self):
  r=self.go();self.assertFalse(r['application']['publication_ready']);self.assertTrue(r['application']['render_required'])
 def test_cli_full_workflow(self):
  tool=Path(f.__file__);plan=self.dir/'plan.json';report=self.dir/'report.json'
  for cmd in [[sys.executable,str(tool),'plan',str(self.src),'--output',str(plan)],[sys.executable,str(tool),'apply',str(self.src),'--plan',str(plan),'--output',str(self.out),'--report',str(report)]]:
   result=subprocess.run(cmd,text=True,capture_output=True);self.assertEqual(result.returncode,0,result.stderr)
  self.assertTrue(json.loads(report.read_text())['application']['stored_text_and_hyperlinks_preserved'])
 def test_import_local_drawio_is_staged(self):
  svg=b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><path d="M0 0L10 10"/></svg>'
  lib=self.dir/'library.xml';lib.write_text('<mxlibrary>'+json.dumps([{'title':'An Icon','w':10,'h':10,'data':'data:image/svg+xml;base64,'+base64.b64encode(svg).decode()}])+'</mxlibrary>')
  out=self.dir/'imported';r=ba.import_library(lib,out,'Synthetic local fixture');self.assertEqual(r['imported'],1);self.assertFalse(json.loads((out/'asset-catalog.staged.json').read_text())['assets'][0]['approved'])
 def test_svg_external_link_rejected(self):
  with self.assertRaises(ValueError):ba.safe_svg(b'<svg xmlns="http://www.w3.org/2000/svg"><use href="https://example.invalid/x"/></svg>')
 def test_svg_script_rejected(self):
  with self.assertRaises(ValueError):ba.safe_svg(b'<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>')
 def test_svg_foreignobject_rejected(self):
  with self.assertRaises(ValueError):ba.safe_svg(b'<svg xmlns="http://www.w3.org/2000/svg"><foreignObject/></svg>')
 def test_svg_normal_style_accepted(self):
  ba.safe_svg(b'<svg xmlns="http://www.w3.org/2000/svg"><style>.a{fill:#DA291C}</style><path class="a" d="M0 0"/></svg>')
 def test_tokens_contain_required_2026_additions(self):
  t=f.jd(f.ROOT/'assets/brand-tokens.json')['colours'];self.assertEqual(t['yellow'],'FFB900');self.assertEqual(t['ot_grey'],'75787B');self.assertEqual(t['red'],'DA291C')
 def test_catalog_hashes_and_no_font_payloads(self):
  for a in f.catalog()['assets']:f.asset_for(a['id'])
  self.assertFalse([p for p in f.ROOT.rglob('*') if p.suffix.lower() in ('.ttf','.otf','.woff','.woff2','.fntdata')])

 def test_duplicate_archive_entry_rejected(self):
  with zipfile.ZipFile(self.src,'a') as z:z.writestr('ppt/media/shared.png',b'duplicate')
  with self.assertRaises(f.BrandError):f.Package(self.src)
 def test_symlink_archive_entry_rejected(self):
  with zipfile.ZipFile(self.src,'a') as z:
   zi=zipfile.ZipInfo('symlink');zi.create_system=3;zi.external_attr=(0o120777<<16);z.writestr(zi,b'/etc/passwd')
  with self.assertRaises(f.BrandError):f.Package(self.src)
 def test_grouped_picture_blocked(self):
  d=self.pkg.doc('ppt/slides/slide1.xml');pic=f.shape_by_id(d,3);g=f.new(d,'p','grpSp');pic.parentNode.replaceChild(g,pic);g.appendChild(pic);self.pkg.save_doc('ppt/slides/slide1.xml')
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'secure-networking'}]
  with self.assertRaises(f.BrandError):self.go()
 def test_rotated_picture_blocked(self):
  p=f.shape_by_id(self.pkg.doc('ppt/slides/slide1.xml'),3);f.one(p,'xfrm').setAttribute('rot','5400000');self.pkg.save_doc('ppt/slides/slide1.xml')
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'secure-networking'}]
  with self.assertRaises(f.BrandError):self.go()
 def test_picture_mask_is_reset_to_rectangle(self):
  p=f.shape_by_id(self.pkg.doc('ppt/slides/slide1.xml'),3);f.one(p,'prstGeom').setAttribute('prst','ellipse');self.pkg.save_doc('ppt/slides/slide1.xml')
  self.plan['replace_images']=[{'slide_part':'ppt/slides/slide1.xml','shape_id':3,'asset_id':'secure-networking'}];self.go()
  self.assertEqual(f.one(f.shape_by_id(f.Package(self.out).doc('ppt/slides/slide1.xml'),3),'prstGeom').getAttribute('prst'),'rect')
 def test_chart_values_and_formula_preserved(self):
  c='http://schemas.openxmlformats.org/drawingml/2006/chart';self.pkg.data['ppt/charts/chart1.xml']=f'<c:chartSpace xmlns:c="{c}" xmlns:a="{A}"><c:numRef><c:f>Sheet1!A1</c:f><c:numCache><c:pt idx="0"><c:v>123.45</c:v></c:pt></c:numCache></c:numRef><a:defRPr><a:latin typeface="Calibri"/></a:defRPr></c:chartSpace>'.encode()
  before=f.fingerprint(self.pkg);self.go();after=f.fingerprint(f.Package(self.out));self.assertEqual(before['chart_math_values'],after['chart_math_values'])
 def test_svg_internal_gradient_accepted(self):
  ba.safe_svg(b'<svg xmlns="http://www.w3.org/2000/svg"><path style="fill:url(#local)" d="M0 0"/></svg>')

if __name__=='__main__':unittest.main(verbosity=2)
