#!/usr/bin/env node
/* Data-driven Fortinet dense architecture explainer.
 * Runtime: Node.js 18+ and locally installed pptxgenjs. No model/network calls.
 * For ChatGPT generation set FORTINET_PPTX_HELPERS to the provided Slides helper
 * module. Outside ChatGPT a small local geometry fallback is used, not a copy of
 * the platform's helper library. Text and connectors stay editable; images do not.
 */
'use strict';
const fs=require('fs'),path=require('path'),crypto=require('crypto');
const pptxgen=require('pptxgenjs');
const ROOT=path.resolve(__dirname,'..');
const tokens=JSON.parse(fs.readFileSync(path.join(ROOT,'assets','brand-tokens.json')));
const catalog=JSON.parse(fs.readFileSync(path.join(ROOT,'assets','asset-catalog.json'))).assets;
const argv=process.argv.slice(2);
function arg(k){let i=argv.indexOf(k);return i<0?null:argv[i+1];}
const input=arg('--data'),output=arg('--output');
if(!input||!output){console.error('Usage: node scripts/build_connected_architecture.cjs --data examples/connected-architecture.json --output output.pptx');process.exit(2);}
if(fs.existsSync(output)){console.error('Refusing to overwrite output');process.exit(2);}
const data=JSON.parse(fs.readFileSync(input,'utf8'));
function check(ok,msg){if(!ok)throw new Error(msg);}
check(data.cards?.length===6,'This template requires six cards');
check(data.nodes?.length===8,'This template requires eight environmental nodes');
check(data.roadmap?.length===4,'This template requires four roadmap steps');
for(const c of data.cards)check(c.rows?.length>=1&&c.rows.length<=4,'Cards require one to four rows');
const helper=process.env.FORTINET_PPTX_HELPERS?require(process.env.FORTINET_PPTX_HELPERS):null;
const pptx=new pptxgen();pptx.layout='LAYOUT_WIDE';pptx.author='Fortinet';pptx.company='Fortinet';pptx.title=data.title; pptx.subject='Conceptual architecture';pptx.lang=data.locale||'en-AU';
pptx.theme={headFontFace:'Inter',bodyFontFace:'Inter',lang:data.locale||'en-AU',themeColors:Object.values(tokens.theme)};
const s=pptx.addSlide(),C=tokens.colours,T=pptx.ShapeType;
s.background={color:C.white};let manifest=[];
function shape(type,o,name,group){s.addShape(type,{...o,objectName:name});manifest.push({type:'shape',name,group,...o});}
function tx(t,x,y,w,h,size=11,opts={}){s.addText(t,{x,y,w,h,fontFace:'Inter',fontSize:size,color:C.black,margin:0,valign:'mid',breakLine:false,...opts});manifest.push({type:'text',name:opts.objectName||String(t),x,y,w,h,fontSize:size});}
function image(id,x,y,w,h,name){
 const a=catalog.find(a=>a.id===id);check(a&&a.approved,'Unknown or unapproved asset: '+id);
 const file=path.join(ROOT,a.path);check(crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex')===a.sha256,'Asset hash mismatch: '+id);
 let b;if(helper)b=helper.imageSizingContain(file,x,y,w,h);else{let sc=Math.min(w/a.width_px,h/a.height_px);let wi=a.width_px*sc,hi=a.height_px*sc;b={x:x+(w-wi)/2,y:y+(h-hi)/2,w:wi,h:hi};}
 s.addImage({path:file,...b,objectName:name||id,altText:(name||a.label)+'; '+a.provenance.source});manifest.push({type:'image',asset_id:id,name:name||id,...b});
}
function line(x1,y1,x2,y2,name){shape(T.line,{x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color:C.red,width:1.1}},name,'connector');}
// Preserve original transparent padding. Wordmark clearance is a reserved region.
image('fortinet-logo-rgb-black-red',10.78,.12,2.20,.70,'Official Fortinet wordmark');
const parts=[{text:data.title_prefix,options:{bold:true}},{text:data.title_accent,options:{bold:true,color:C.red}},{text:data.title_suffix,options:{bold:true}}];
tx(parts,.43,.28,10.02,.48,28.5,{bold:true,objectName:'Title'});
tx(data.subtitle,.45,.90,12.37,.25,12.2,{color:C.dark_grey,objectName:'Subtitle'});
function rich(text){let i=text.indexOf(':');return i<0?text:[{text:text.slice(0,i+1),options:{bold:true}},{text:text.slice(i+1),options:{bold:false}}];}
const ys=[1.37,3.12,4.87];
data.cards.forEach((c,i)=>{
 const x=i<3?.43:9.20,y=ys[i%3],w=i<3?3.50:3.70,h=1.67,name='Card '+(i+1);
 shape(T.roundRect,{x,y,w,h,radius:.11,rectRadius:.11,line:{color:C.light_grey,transparency:100},fill:{color:C.light_grey}},name+' panel',name);
 tx(String(i+1).padStart(2,'0'),x+.17,y+.17,.31,.26,14.3,{bold:true,color:C.red,objectName:name+' number'});
 tx(c.title,x+.59,y+.15,w-1.14,.31,c.title_font_size||13,{bold:true,objectName:name+' title'});
 image(c.icon,x+w-.45,y+.155,.29,.30,name+' icon');let cursor=y+.49;
 c.rows.forEach((r,j)=>{const h=r.height||.21;tx(rich(r.text),x+.18,cursor,w-.36,h,c.body_font_size||10.75,{valign:'top',objectName:name+' item '+(j+1)});cursor+=h+(r.gap??.04);});
 check(cursor<=y+1.67+.01,name+' content boxes exceed reserved space; shorten copy or split the slide, do not auto-shrink');
});
const start=4.16,nw=1.09,nh=.99,xs=[0,1,2,3].map(i=>start+i*1.22),top=1.58,bottom=5.32,cx=6.53;
line(xs[0]+nw/2,2.86,xs[3]+nw/2,2.86,'Upper bus');line(cx,2.86,cx,3.15,'Upper hub link');
line(xs[0]+nw/2,4.97,xs[3]+nw/2,4.97,'Lower bus');line(cx,4.73,cx,4.97,'Lower hub link');
for(let i=0;i<4;i++){line(xs[i]+nw/2,top+nh,xs[i]+nw/2,2.86,'Top link '+i);line(xs[i]+nw/2,4.97,xs[i]+nw/2,bottom,'Bottom link '+i);}
shape(T.roundRect,{x:5.08,y:3.15,w:2.90,h:1.58,radius:.10,rectRadius:.10,line:{color:C.red,width:1.2},fill:{color:C.white}},'Hub','hub');
image('fortinet-logomark-rgb-red',6.15,3.19,.77,.61,'Official Fortinet Grid');
tx(data.hub.eyebrow,5.23,3.89,2.60,.18,10,{color:C.dark_grey,bold:true,align:'center',objectName:'Hub eyebrow'});
tx(data.hub.title,5.21,4.11,2.64,.26,17.5,{color:C.red,bold:true,align:'center',objectName:'Hub title'});
tx(data.hub.subtitle,5.21,4.43,2.64,.19,11.2,{bold:true,align:'center',objectName:'Hub subtitle'});
data.nodes.forEach((n,i)=>{const x=xs[i%4],y=i<4?top:bottom;shape(T.roundRect,{x,y,w:nw,h:nh,radius:.07,rectRadius:.07,line:{color:C.grey,width:.65},fill:{color:C.white}},'Node '+i,'node');image(n.icon,x+.275,y+.11,.54,.48,n.alt||n.label);tx(n.label,x+.035,y+.64,nw-.07,.27,9.3,{bold:true,align:'center',objectName:'Node label '+i});});
// A native chevron plus a narrow same-colour patch forms the first square-ended
// segment. This intentional overlay is limited to the empty left edge, not text.
let ry=6.70,rh=.40;
shape(T.chevron,{x:.43,y:ry,w:3.66,h:rh,line:{color:C.red,width:0},fill:{color:C.red}},'Roadmap title','roadmap');
shape(T.rect,{x:.43,y:ry,w:.21,h:rh,line:{color:C.red,width:0},fill:{color:C.red}},'Roadmap left-edge patch','roadmap');
tx(data.roadmap_title,.67,ry+.115,3.11,.16,10.6,{bold:true,color:C.white,objectName:'Roadmap title text'});
[4.12,6.33,8.54,10.75].forEach((x,i)=>{shape(T.chevron,{x,y:ry,w:2.15,h:rh,line:{color:C.red,width:.75},fill:{color:C.red_t1}},'Roadmap step '+i,'roadmap');tx(data.roadmap[i],x+.20,ry+.12,1.74,.17,i===1?9.4:9.8,{align:'center',objectName:'Roadmap label '+i});});
tx(data.footer,.44,7.18,12.42,.18,11.2,{color:C.dark_grey,align:'center',objectName:'Footer'});
const sources=[...new Set(catalog.map(a=>a.provenance.source))];
s.addNotes('Conceptual architecture example. Source content is a user-supplied concept, not a verified product capability statement. The Inter presentation profile is explicitly user approved; the 2024 guideline retains a legacy Arial exception. Some labels are below 10 pt because this is a dense, downloadable explainer, not a default large-room presentation template. Split into multiple slides for projection.\n\n'+(data.notes||'')+'\n\n[Sources]\n'+sources.join('\n')+'\nFortinet Editorial Style Guide, September 2026.\nArtwork sources and crop lineage: assets/asset-catalog.json.\n[/Sources]');
if(helper){helper.warnIfSlideHasOverlaps(s,pptx);helper.warnIfSlideElementsOutOfBounds(s,pptx);}
for(const o of manifest)check(o.x>=0&&o.y>=0&&o.x+o.w<=40/3+.005&&o.y+o.h<=7.505,'Out-of-bounds object: '+o.name);
fs.mkdirSync(path.dirname(path.resolve(output)),{recursive:true});
pptx.writeFile({fileName:output}).then(()=>{fs.writeFileSync(output.replace(/\.pptx$/i,'.layout.json'),JSON.stringify({template:'dense-connected-architecture',version:'2.0.0',font:'Inter',density_warning:'Several node/roadmap labels are below 10 pt. Split for projection.',objects:manifest},null,2));console.log(JSON.stringify({output,render_required:true}));}).catch(e=>{console.error(e.message);process.exitCode=2;});
