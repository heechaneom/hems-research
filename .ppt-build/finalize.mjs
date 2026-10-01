import fs from 'node:fs/promises';
import path from 'node:path';
import {Presentation, PresentationFile} from '@oai/artifact-tool';
import {finalizePresentation} from 'file:///C:/Users/KBCARD/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations/container_tools/artifact_tool_utils.mjs';

const W='D:/hems-search', B=W+'/.ppt-build';
const SK='C:/Users/KBCARD/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations';
const PY='C:/Users/KBCARD/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const P=Presentation.create({slideSize:{width:1280,height:720}});
const C={navy:'#102A43',ink:'#17324D',muted:'#526777',teal:'#007F78',paper:'#F5F6F2',white:'#FFFFFF',gold:'#E5BB66'};
const FONT='Malgun Gothic';
const S={
 meeting:['공유 대화 · 기술 통합 방안 논의','https://chatgpt.com/share/6abc9e1d-4f6c-83ee-b30e-1eafc9fbbf7c'],
 de:['Bundesnetzagentur · Dynamische Stromtarife','https://www.bundesnetzagentur.de/DE/Vportal/Energie/Vertragsarten/DynStromtarife/start.html'],
 de14:['Bundesnetzagentur · §14a 제어가능 부하','https://www.bundesnetzagentur.de/DE/Vportal/Energie/SteuerbareVBE/artikel.html?nn=779422'],
 uk:['Ofgem · 2025 유연성 소비자 조사, 2026.05.22','https://www.ofgem.gov.uk/research/tracking-energy-consumers-use-low-carbon-and-flexible-products-and-services-2025'],
 fr:['Enedis · 비첨두 시간대 개편, 2025.10.15','https://www.enedis.fr/presse/enedis-met-en-oeuvre-la-reforme-des-heures-creuses-pour-accompagner-les-nouveaux-usages-lies'],
 nl:['네덜란드 정부 · 2027년 상계제도 종료','https://www.rijksoverheid.nl/themas/klimaat-milieu-en-natuur/energie-thuis/salderingsregeling'],
 au:['호주 정부 · Solar PV and batteries','https://www.energy.gov.au/households/solar-pv-and-batteries'],
 ca:['CPUC · Net Billing Tariff','https://www.cpuc.ca.gov/industries-and-topics/electrical-energy/demand-side-management/customer-generation/nem-revisit/net-billing-tariff'],
 vpp:['미국 DOE · Virtual Power Plants Projects','https://www.energy.gov/edf/virtual-power-plants-projects'],
 uae:['DEWA · Cooling','https://www.dewa.gov.ae/en/consumer/sustainability/sustainability-and-conservation/cooling'],
 kr:['통계청 · 2024 인구주택총조사, 표 36','https://www.kostat.go.kr/boardDownload.es?bid=203&list_no=437767&seq=3'],
 za:['Eskom · 전력계통 현황, 2026.09.25','https://www.eskom.co.za/eskom-achieves-weekly-eaf-of-70-and-year-to-date-eaf-of-68-06-emissions-reduction-target-surpassed-diesel-costs-fall-by-r4-8-billion-year-on-year-and-load-reduction-further-reduced-from-1-69-million/'],
 sma:['SMA · Sunny Home Manager 2.0','https://www.sma.de/en/products/energy-management/sunny-home-manager'],
 tesla:['Tesla · Savings Mode','https://www.tesla.com/en_gb/support/energy/powerwall/mobile-app/savings-mode'],
 oct:['Octopus · Heat pump FAQ','https://octopus.energy/heat-pump-FAQ/'],
 sch:['Schneider · Residential products','https://www.se.com/us/en/work/products/residential-and-small-business/']
};
let pages=[];const tableOwners=[];
function txt(sl,t,x,y,w,h,size=26,color=C.ink,bold=false){const z=sl.shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});z.text=t;z.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none'};return z;}
function slide(title,subtitle='',refs=[],dark=false){let s=P.slides.add();s.background.fill=dark?C.navy:C.paper;let n=pages.length+1;pages.push({s,title,refs});txt(s,title,62,45,1156,94,42,dark?C.white:C.ink,true);if(subtitle)txt(s,subtitle,64,143,1150,55,22,dark?'#C5D8E7':C.muted);txt(s,String(n).padStart(2,'0'),1160,675,64,25,15,dark?'#9BB7CA':C.muted);txt(s,'ENERGY PACK  /  신사업·신기술 검토  /  2026.09',64,675,1030,25,14,dark?'#9BB7CA':C.muted);s.speakerNotes.textFrame.setText('조사 기준일: 2026-09-30. 외부 사실과 작성자의 사업 제안/가설을 구분함. 목표 수치는 실적이나 확정 사업계획이 아님.\n'+refs.map(k=>S[k].join('\n')).join('\n\n'));return s;}
function blocks(s,arr,{top=224,gap=140,left=66,width=1135,dark=false}={}){arr.forEach((a,i)=>{txt(s,a[0],left,top+i*gap,width,42,29,dark?C.gold:C.teal,true);txt(s,a[1],left,top+48+i*gap,width,80,26,dark?C.white:C.ink);});}
function cols(s,arr){arr.forEach((a,i)=>{let x=64+i*592;txt(s,a[0],x,219,548,58,31,C.teal,true);txt(s,a[1],x,291,545,315,27);});}
function table(s,values,widths,{top=216,height=380,size=24}={}){let t=s.tables.add({rows:values.length,columns:values[0].length,left:64,top,width:1152,height,values,columnWidths:widths});t.borders.assign({fill:'#D8E0E2',width:0.6,style:'solid'});for(let r=0;r<values.length;r++)for(let c=0;c<values[0].length;c++){let cell=t.getCell(r,c);cell.fill=r===0?C.navy:(r%2?C.white:'#EAF0EE');cell.text.style={typeface:FONT,fontSize:size,bold:r===0,color:r===0?C.white:C.ink};}tableOwners.push(pages.length);return t;}
function note(s,t,y=622){txt(s,t,66,y,1144,44,18,C.muted);}


tableOwners.push(3,4,5,11,12,16,18,19,20,21,22,23,24,26);
const finalPath=W+'/output/Energy_Pack_국가별_PainPoint_미래신기술_20260930_최종.pptx';
const result=await finalizePresentation({workspaceDir:W,candidatePath:B+'/candidate.pptx',finalPath,pythonExecutable:PY,integrityValidatorPath:SK+'/container_tools/inspect_presentation_package_integrity.py',layoutValidatorPath:SK+'/container_tools/inspect_presentation_layout_geometry.py',layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:tableOwners,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:B+'/validation-final.json'});
console.log(JSON.stringify(result));




