import fs from 'node:fs/promises';
import {Presentation, PresentationFile} from '@oai/artifact-tool';
import {finalizePresentation} from 'file:///C:/Users/KBCARD/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations/container_tools/artifact_tool_utils.mjs';

const W='D:/hems-search',B=W+'/.gateway-research-build';
const SK='C:/Users/KBCARD/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations';
const PY='C:/Users/KBCARD/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const OUT=W+'/output/Energy_Pack_Gateway_Strategy_2027-2030_Research_20260930_최종.pptx';
const FONT='Malgun Gothic';
const C={ink:'#142E42',muted:'#5A6D77',teal:'#007B74',gold:'#D8B26A',paper:'#FAFBF8',navy:'#112A3A',white:'#FFFFFF'};
const P=Presentation.create({slideSize:{width:1280,height:720}});
const src={
 ref:['R00','제공 HTML','Energy Pack Gateway Technology 2027–2030','D:/hems-search/output/Energy_Pack_Gateway_Technology_2027-2030.html'],
 de:['R01','BNetzA','§14a 제어가능 설비 적용 안내','https://www.bundesnetzagentur.de/DE/Vportal/Energie/SteuerbareVBE/artikel.html'],
 uk:['R02','NESO','Demand Flexibility Service explained, 2026 개편','https://www.neso.energy/industry-information/balancing-services/demand-flexibility-service/demand-flexibility-service-explained'],
 stack:['R03','NESO','DFS Stacking List v2, 2026.04','https://www.neso.energy/document/380211/download'],
 fr:['R04','Enedis','Heures creuses 개편, 2025.10.15','https://www.enedis.fr/presse/enedis-met-en-oeuvre-la-reforme-des-heures-creuses-pour-accompagner-les-nouveaux-usages-lies'],
 nl:['R05','네덜란드 정부','Salderingsregeling stopt in 2027','https://www.rijksoverheid.nl/themas/klimaat-milieu-en-natuur/energie-thuis/salderingsregeling'],
 it:['R06','ARERA','Potenza impegnata 및 계량기 용량 안내','https://www.arera.it/bolletta/glossario-dei-termini/dettaglio/potenza-impegnata'],
 it2:['R07','ARERA','Potenza del contatore','https://www.arera.it/en/consumatori/potenza-del-contatore'],
 es:['R08','CNMC','2023년 말 가계조사, 2024.06.28 발표','https://www.cnmc.es/prensa/panel-hogares-servicios-electricidad-gas-20240628'],
 ca:['R09','CPUC','Net Energy Metering and Net Billing','https://www.cpuc.ca.gov/industries-and-topics/electrical-energy/demand-side-management/customer-generation/net-energy-metering-and-net-billing'],
 tx:['R10','ERCOT','ADER Pilot, Phase 3.3, 2026.06.03','https://www.ercot.com/mktrules/pilots/ader'],
 irs:['R11','IRS','Residential Clean Energy Credit','https://www.irs.gov/credits-deductions/residential-clean-energy-credit'],
 ae:['R12','DEWA','Slab Tariff, 연료조정요금 2026.09','https://www.dewa.gov.ae/en/consumer/billing/slab-tariff'],
 sa:['R13','SERA','Electricity Service Guide, 요금표 p.67','https://www.sera.gov.sa/-/media/4a51a8ff4b5a45609033cfb7fea9d644.ashx'],
 qa:['R14','KAHRAMAA','BeSolar net billing 안내','https://km.qa/RenewableEnergy/Pages/Besolar.aspx'],
 gridx:['R15','gridX','Optimized Heat Pump Operation, 2026.06.02','https://support.gridx.ai/en/articles/590093-dynamic-tariff-optimization-optimized-heat-pump-operation'],
 hb:['R16','1KOMMA5°','Heartbeat 호환 목록, 2026.08.25','https://support.1komma5.com/hc/de/articles/29833157635356-Heartbeat-Kompatibilit%C3%A4tsliste'],
 kraken:['R17','Kraken','Electrifying domestic heating, 2025.08.21','https://engineering.kraken.tech/news/2025/08/21/electrifying-domestic-heating.html'],
 enphase:['R18','Enphase','IQ Energy Router 독일·오스트리아 출시, 2023.06','https://investor.enphase.com/news-releases/news-release-details/enphase-energy-launches-new-home-energy-management-solution/'],
 enphase2:['R19','Enphase','Daikin Altherma 3 설정 가이드','https://enphase.com/en-de/download/heat-pump-daikin-altherma-3-configuration-guide-iq-energy-router'],
 span:['R20','SPAN','PowerUp 및 배전반 에너지 제어','https://www.span.io/'],
 cd:['R21','ConnectDER','Solar MSA v5.2 설치 매뉴얼','https://connectder.com/wp-content/uploads/2023/08/ConnectDER_Solar_MSA_v5.2_Installation-Manual-v1.0f.pdf'],
 ul:['R22','UL Solutions','Power Control Systems for DER','https://www.ul.com/services/power-control-systems-distributed-energy-resource-systems'],
 g100:['R23','ENA','G100 Issue 2 Amendment 2, 2023','https://www.energynetworks.org/assets/images/ENA_EREC_G100_Issue_2_Amendment_2_%282023%29.pdf'],
 heat:['R24','Nesta·CNZ','HeatFlex UK, 2024 실증','https://www.nesta.org.uk/project/heat-flex/'],
 heat2:['R25','Nesta','HeatFlex 기술·실증 보고서','https://media.nesta.org.uk/documents/HeatFlex_-_the_untapped_potential_of_heat_pump_flexibility.pdf'],
 eebus:['R26','EEBUS','Solutions & Use Cases','https://www.eebus.org/solutions-use-cases/'],
 eebus2:['R27','EEBUS','Limitation of Power Consumption v1.0','https://www.eebus.org/wp-content/uploads/2023/04/EEBus_UC_TS_LimitationOfPowerConsumption_V1.0.0_public.pdf'],
 matter:['R28','CSA','Matter 1.5 에너지 관리 확장, 2025.11.20','https://csa-iot.org/newsroom/matter-1-5-introduces-cameras-closures-and-enhanced-energy-management-capabilities/'],
 cra:['R29','EU 집행위','Cyber Resilience Act 적용 일정','https://digital-strategy.ec.europa.eu/en/policies/cra-summary'],
 patent:['R30','특허 공개문헌','US10191102B2, CT 극성 자동 보정','https://patents.google.com/patent/US10191102B2/en'],
};
let pages=[],tableOwners=[];
function tx(s,text,x,y,w,h,size=27,color=C.ink,bold=false){let t=s.shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});t.text=text;t.text.style={typeface:FONT,fontSize:size,color,bold,autoFit:'none'};return t;}
function add(title,sub='',refs=[],{dark=false,foot='',notes=''}={}){let s=P.slides.add();s.background.fill=dark?C.navy:C.paper;let n=pages.length+1;pages.push({s,title,refs});tx(s,title,64,43,1150,102,42,dark?C.white:C.ink,true);if(sub)tx(s,sub,66,151,1146,62,23,dark?'#BDD1D8':C.muted);tx(s,String(n).padStart(2,'0'),1158,676,64,24,16,dark?'#BDD1D8':C.muted);tx(s,'Energy Pack 2027–2030  /  2026.09.30',66,676,1050,24,15,dark?'#BDD1D8':C.muted);let refsText=refs.length?'근거  '+refs.map(k=>src[k][0]+' '+src[k][1]).join(' · '):'사업 제안 · 수치가 있는 경우 별도 표시한 가정 또는 목표';if(foot)refsText=foot+'\n'+refsText;tx(s,refsText,66,626,1150,48,15,dark?'#BDD1D8':C.muted);s.speakerNotes.textFrame.setText('조사 기준일: 2026-09-30.\n공개 1차 자료의 확인 사실과 작성자의 사업 제안·가설·목표를 구분했다. 내부 성능·원가·채널 역량은 확인되지 않았다.\n'+notes+'\n\n'+refs.map(k=>src[k].join('\n')).join('\n\n'));return s;}
function duo(s,a,b,{top=230,size=27,height=330,dark=false}={}){for(let [i,v] of [a,b].entries()){let x=66+i*591;tx(s,v[0],x,top,544,51,30,dark?C.gold:C.teal,true);tx(s,v[1],x,top+65,544,height,size,dark?C.white:C.ink);}}
function lines(s,arr,{top=226,gap=126,dark=false}={}){arr.forEach((v,i)=>{tx(s,v[0],67,top+i*gap,1124,42,29,dark?C.gold:C.teal,true);tx(s,v[1],69,top+49+i*gap,1120,75,27,dark?C.white:C.ink);});}
function table(s,values,widths,{top=223,height=371,size=24}={}){let t=s.tables.add({rows:values.length,columns:values[0].length,left:66,top,width:1148,height,values,columnWidths:widths});t.borders.assign({fill:'#DDE4E2',width:0.6,style:'solid'});for(let r=0;r<values.length;r++)for(let c=0;c<values[0].length;c++){let z=t.getCell(r,c);z.fill=r===0?C.navy:(r%2?C.white:'#EEF3EF');z.text.style={typeface:FONT,fontSize:size,bold:r===0,color:r===0?C.white:C.ink};}tableOwners.push(pages.length);return t;}

// 01
{
let s=add('Energy Pack','', ['ref'],{dark:true,foot:'국가별 고객 문제와 기술 개발·사업화 우선순위'});
tx(s,'미래 먹거리를 만드는\n관문기술과 진입시장',68,234,1110,169,58,C.white,true);
tx(s,'2027–2030',72,431,1050,69,46,C.gold,true);
tx(s,'국가별 Pain Point · 경쟁 제품 검증 · R&D 실행안',72,543,1120,46,27,'#BDD1D8');
}
// 02
{
let s=add('2027년은 두 과제에 집중하는 편이 타당하다','독일의 설치 검증과 영국의 열 유연성을 먼저 검증하고, 동일 기술을 네덜란드로 확장',[],{notes:'경영 판단을 위한 우선순위 제안이다. 시장 규모에 대한 정량 점수나 확보된 수주를 의미하지 않는다. 두 과제의 공통 기반은 제어능력 확인, 계측, 현장 시험, 기기별 제약 데이터이다.'});
lines(s,[['직접 개발  GT1 설치 검증 자동화','설치업체가 반복하는 연결·설정·시험 작업을 줄이고, 유료 도입 여부를 확인한다.'],['직접 개발  GT3 쾌적 제약을 지키는 열 제어','집별로 실제 제공 가능한 전력 조절 범위를 추정하고, 단순 타이머 대비 편익을 검증한다.'],['제휴 우선  GT2 전력 허브와 GT4 시장 연계','전력 보호·인증 하드웨어와 시장 등록·정산은 파트너를 활용한다.']]);
}
// 03
{
let s=add('기존 논리에서 수정해야 할 다섯 가지','4대 관문기술 틀은 유지하되, 사업 범위와 차별성의 근거를 좁혔다',['fr','tx','hb','gridx','ul','stack']);
table(s,[['기존 자료의 주요 전제','추가 확인 결과','전략에 반영할 변화'],['프랑스 급탕 스케줄 대규모 재설정','Linky 접점 연결 기기는 자동 추종','고정 타이머·혼합 설비만 선별'],['다른 브랜드 통합은 경쟁사 공백','복수 제조사 연동 제품이 이미 존재','연결 후 실제 제어·시험에 집중'],['텍사스 ADER의 전면 시장화','2026.06 Phase 3.3도 Pilot','등록 파트너 확보 전 수익 제외'],['소프트웨어가 증설 문제 해결','PCS 인증·현장 조건에 따라 다름','인증 제품과 경제성 비교'],['여러 프로그램 수익을 동시 합산','시간·용량·방향별 중복 규칙 존재','허용 계약만 편익에 반영']], [342,403,403],{height:382,size:23});
}
// 04
{
let s=add('시장 우선순위는 고객 접근성과 실증 조건으로 결정','국가 수를 세어 수요를 추정하지 않고, 구매자·제어권·계량 접근 가능성을 함께 판단',[]);
table(s,[['진입 순서 제안','첫 고객과 문제','진입 조건'],['1  독일','설치업체·OEM의 설정·시험 부담','지원 기기 목록, 현장 30건 접근'],['1  영국','소매사·집합사업자의 열부하 운전','계량 데이터·기기 제어권·실증 가구'],['2  네덜란드','PV 가구의 2027년 이후 자가소비','HP·급탕탱크 보유, 수출 계약 확인'],['선별  프랑스·스페인·이탈리아','고정 타이머·계약전력·혼합 설비','저비용 대안 대비 추가 가치 확인'],['파트너 동반  미국·걸프','전력 증설·백업 또는 냉방 운전','인증·채널·요금 부담자 확보']], [289,462,397],{height:384,size:24});
}
// 05
{
let s=add('독일: 설치 완료를 검증하는 비용을 줄인다','첫 구매자 후보는 가정 소비자보다 설치업체·Energy Pack OEM',['de','eebus'],{foot:'§14a는 대상 설비의 수전 제어다. 일반 가전과 PV 역송 전체를 한꺼번에 제한하는 제도가 아니다.'});
duo(s,['확인한 제도와 고객 문제','2024년 이후 신규 설치 등 조건에 따라\n히트펌프·EV 충전·저장장치 등이 적용\n\nEMS가 여러 대상 설비를 조정할 수 있음\n\n현장 문제 가설\n계측 방향·제어권·통신·설정이 맞는지\n기기별로 확인하는 시간이 길다.'],['판매할 모듈과 검증 방법','기기 연결 후 명령·응답 시험 수행\n\n지원 조합별 설치 완료 보고서 생성\n\n수익 방식: 설치 건별 또는 OEM 탑재료\n\n측정: 작업시간, 재방문, 실패 원인\n비교: 현재 설치 앱과 기존 HEMS']);
}
// 06
{
let s=add('영국: 열부하의 양방향 유연성을 공급한다','소매사·집합사업자가 고객 확보와 시장 참여를 담당하는 B2B2C 모델',['uk','heat','stack'],{foot:'제도 참여가 수익을 보장하지 않는다. 대상 가구·이벤트·보상은 공급자 계약에 따라 달라진다.'});
duo(s,['2026년의 기회','NESO DFS는 2026.04.14부터\n사용 감소뿐 아니라 증가도 보상 가능\n\n스마트미터와 등록 사업자를 통해 참여\n대상 지역: 잉글랜드·스코틀랜드·웨일스\n\n열 유연성은 이미 현장 실증이 존재\n새 과제는 집별 제공량과 신뢰도다.'],['제안 상품','저렴하거나 계통이 요구하는 시간에\n급탕·예열하고 피크 시간 사용을 조절\n\n집별 온도 제약과 회복 부하까지 관리\n\n소매사에는 가용 kW와 불확실성 제공\n고객에게는 요금·온도 결과 공개']);
}
// 07
{
let s=add('네덜란드는 상품화, 프랑스는 대상 선별이 핵심','같은 급탕 제어라도 제도 변화가 만들어내는 추가 가치는 다르다',['nl','fr']);
duo(s,['네덜란드  PV·HP 보유 가구','2027.01.01 상계제도 종료\n역송 보상은 계속 존재\n\n기회: 잉여 발전을 열로 활용하되\n역송 보상·열손실을 함께 계산\n\n채널: PV·HP 설치업체\n판단: 탱크·배터리의 추가 구입 없이도\n운전 개선으로 편익이 남는가?'],['프랑스  미연동 설비를 선별','비첨두 시간 변경은 2025–2027 순차 적용\n\nLinky 접점 연결 급탕기는 자동 추종\n1,100만 변경 고객 모두가 추가 시장은 아님\n\n기회: 고정 타이머, PV·HP 혼합 제어\n채널: 설비 유지보수 사업자\n판단: 간단한 재설정으로 해결되는가?'],{size:26});
}
// 08
{
let s=add('남유럽: 용량 부족과 과다 계약을 구분한다','증설 회피형 허브는 실제로 비싼 공사가 필요한 집에 한해 검토',['it','it2','es'],{foot:'스페인 수치는 2023년 말 조사다. 현재 전체 시장 규모나 신규 전기화 가구의 상황으로 확대 해석하지 않는다.'});
duo(s,['이탈리아  동시 사용 제약','일반 가정의 대표 계약전력은 3 kW\n계량기는 통상 10% 여유를 제공\n\n기회: EV·히트펌프를 기존 가전과 조정\n\n먼저 비교할 대안\n계약용량 변경, EV 단독 부하관리\n\n허브 가격이 대안보다 낮아야 성립'],['스페인  계약 최적화부터','CNMC 조사에서 63% 가구가\n피크 시간 실제 사용보다 높은 용량 계약\n\n모든 집이 용량 부족에 처한 것은 아님\n\n기회: 전기화 예정 가구의 사용 이력으로\n계약 변경과 능동 제어를 함께 제안\n\n재투입·증설 비용까지 비교']);
}
// 09
{
let s=add('미국: 캘리포니아 운전 최적화와 텍사스 참여를 분리','보조금·백업·전력시장 수익은 같은 경제성 항목으로 자동 합산하지 않는다',['ca','tx','irs'],{foot:'IRS §25D는 2025년 이후 사용 개시 설비에 적용되지 않는다. 다른 사업·주정부 제도의 조건은 별도 확인.'});
duo(s,['캘리포니아  NBT 가구','역송 보상은 보통 소매요금보다 낮지만\n일부 늦여름 저녁에는 더 높을 수 있음\n\n제안: 열부하·배터리·백업 잔량 통합 운전\n\n자가소비 극대화만으로 최적을 단정 불가\n가입 연도·요금표별 실제 청구로 검증\n\n채널: 인증 ESS·설치 파트너'],['텍사스  파트너 실증','ERCOT ADER는 Phase 3.3 Pilot\n\n참여에는 등록·계량·텔레메트리와\n관련 사업자의 확인 절차 필요\n\n제안: 현지 집합사업자에 제어 모듈 공급\n\n시장 계약 확보 전 기본 수익은 0으로 설정\n고객 백업 선호를 계약에 반영'],{size:26});
}
// 10
{
let s=add('걸프: 냉방 요금 절감과 수요반응 사업을 구분한다','국가별 요금과 실제 비용 부담자가 다르므로 하나의 GCC 상품으로 묶기 어렵다',['ae','sa','qa']);
table(s,[['지역·확인 사실','구매자와 문제 가설','우선 제안'],['두바이 / DEWA 누진요금\n0.230–0.380 AED/kWh¹','냉방비를 내는 빌라 소유자·운영사\n운전 낭비, 온습도 불만','열·습도 제어와 이상 탐지\n시간 이동보다 총 소비 절감 검증'],['사우디 / 주거용 누진요금\n월 6,000 kWh 경계, 18·30 halala','주택·단지 운영사\n냉방 요금·유지보수 부담','냉방 최적화 실증\n보조·계약 조건별 편익 확인'],['카타르 / BeSolar net billing\n역송 0.237 QAR/kWh','PV 설치 고객·개발사업자\n자가소비와 역송 중 선택','한계 구입요금과 역송가 비교\n경제성이 낮으면 제어 확대 보류']], [370,394,384],{height:373,size:24});
tx(s,'¹ 기본 전력량 요금. 2026.09 연료조정요금 0.060 AED/kWh와 VAT는 별도',69,600,1120,25,17,C.muted);
}
// 11
{
let s=add('다른 브랜드 연동 자체는 이미 경쟁 영역이다','공식 제품·지원 문서에 나온 기능과, 우리가 추가 검증할 차이를 분리',['gridx','hb','kraken','enphase','enphase2']);
table(s,[['제품·플랫폼','공식 문서에서 확인한 기능','차별성 검증 질문'],['gridX','저가·PV 잉여 시 HP 운전 권고\n해당 기능은 직접 전력값 제어가 아님','실제 응답량을 더 정확하게\n추정·확인할 수 있는가?'],['1KOMMA5° Heartbeat','복수 제조사 호환 목록 제공\n일부 HP는 계량 방식 통합','계측 가능과 제어 가능을\n자동으로 구분할 수 있는가?'],['Kraken','열부하 유연성 프로젝트·운영 경험','소매사 플랫폼에 공급할\nOEM 제어 모듈이 필요한가?'],['Enphase IQ Energy Router','독일·오스트리아에서\n선택된 타사 EV·HP 통합','특정 지원 조합에서 설치·시험을\n얼마나 더 줄일 수 있는가?']], [258,462,428],{height:384,size:23});
}
// 12
{
let s=add('전력 허브도 대체 제품과 인증 경로가 존재한다','하드웨어 사업은 독자 개발보다 기존 제품 대비 추가 가치 검증이 먼저다',['span','cd','ul','g100']);
table(s,[['비교 대상','이미 제공하는 역할','우리가 검증할 추가 가치'],['SPAN','회로별 제어·백업 부하 우선순위\n조건에 맞는 전력 증설 회피','온도·급탕·EV 출발시간까지 반영한\n부하 배분이 추가 편익을 주는가?'],['ConnectDER','계량기 소켓 어댑터를 통한\n설비 접속 구조 단순화','현장 공사비가 문제인지\n운영 제약이 문제인지 구분'],['UL PCS / 영국 G100','전력 제한·고장 시 동작을 다루는\n평가·적합성 경로','국가·설비별 시험 범위와\n인증 파트너 비용을 확보']], [272,456,420],{height:352,size:25});
tx(s,'백업 제어·증설 회피가 미개척이라는 전제로 신규 배전반을 개발하면 차별성이 약하다.',69,588,1110,35,23,C.teal,true);
}
// 13
{
let s=add('4대 관문기술의 개발 범위','관문기술: 고객 도입이나 반복 판매를 막는 문제를 해결하는 공통 기술',[]);
table(s,[['관문기술','남길 기술적 핵심','2027년 선택'],['GT1  설치 검증 자동화','제어능력 탐색, 안전한 응답 시험\n기기·계측·정책 버전별 증적','직접 개발\n독일 설치 파트너'],['GT3  쾌적 제약 열 제어','가용 전력·에너지와 불확실성\n응답 확인과 회복 운전','직접 개발\n영국 소매사·HP OEM'],['GT2  전력 제한 허브','계통·회로·상별 제약 안의 부하 배분','인증 HW 제휴\n증설비가 높은 현장 선별'],['GT4  다중 계약·편익 검증','제어권 충돌 처리, 허용 중복 규칙\n계약별 기록과 고객 순편익','공통 기능으로 구현\n시장 정산은 파트너']], [338,470,340],{height:382,size:24});
}
// 14
{
let s=add('GT1: 자동 발견보다 설치 후 응답 시험이 핵심','지원 기기·정상 배선·승인된 시험 조건 안에서 자동화를 시작',['gridx','enphase2','eebus'],{foot:'설치 보고서는 제품·현장 적합성 판단의 증거다. 법정 인증이나 설치자의 책임을 대체하지 않는다.'});
table(s,[['단계','기술 동작','출력'],['1  관찰','기기·펌웨어·프로토콜·계측점 식별','측정 / 권고 / 직접제어 등 능력 분류'],['2  제한된 시험','허용된 작은 명령을 순차 적용\n실제 전력·상·상태 변화 확인','명령 지연, 미응답, 계측 불일치'],['3  대조','기기 보호·제어권·국가 설정 대조\n실패 시 읽기 전용 또는 수동 절차','통과·재시험·현장 점검 구분'],['4  기록','시간·버전·설정·시험 결과를 저장','설치 보고서와 유지보수 기준선']], [226,511,411],{height:380,size:24});
}
// 15
{
let s=add('GT1 경제성: 줄일 수 있는 작업만 계산한다','가상 사례. 전체 설비 공사비나 태양광의 모든 간접비를 절감 대상으로 잡지 않음',[],{foot:'아래는 시장 평균·실적·제안가격이 아니다. 현장 관찰과 견적을 받기 위한 계산 예시다.',notes:'가정: 설정·시험 3h에서1.5h, 시간당 완전원가€60, 재방문확률12%에서5%, 재방문당€180. 절감90+12.6=102.6. 건별사용료€35와추가자재€20 제외 설치업체 잔여편익€47.6. 설치업체 가치이며당사이익아님. 지원비·개발비는 별도 제품수익성모형으로 검토.'});
table(s,[['계산 항목','가정','건당 편익'],['설정·시험 노동','3.0시간에서 1.5시간, €60/시간','€90.0'],['재방문 기대비용','12%에서 5%, 1회 €180','€12.6'],['설치업체 총 절감','작업·재방문 항목 합계','€102.6'],['사용료와 추가 자재','€35 + €20 차감','−€55.0'],['설치업체에 남는 편익','총 절감 − 추가 비용','€47.6']], [400,478,270],{height:375,size:25});
}
// 16
{
let s=add('GT3: “지금 이 집이 얼마나 조절 가능한가”를 계산한다','시각별 가용 전력, 지속시간, 회복 부하를 불확실성과 함께 제공',['heat','gridx'],{foot:'HeatFlex는 43가구·30회 이벤트의 현장 연구다. 모든 기후·주택·기기에 일반화한 수익 증거는 아니다.'});
duo(s,['입력과 연구 대상','실내·외 온도, 온수 온도, 사용 이력\nHP 상태와 전력, 사용자 허용 범위\n\n미지 상태: 단열·열용량·탱크 잔열\n동작 지연: 제상·최소 운전·휴지시간\n\n핵심 연구\n불확실할수록 제공량을 줄이는\n보수적 가용량 추정'],['판매 가능한 출력','가능한 증감 전력과 지속시간\n\n명령의 실제 이행 여부와 원인\n\n이벤트 후 회복 전력과 예상 비용\n\n사용자는 온도 범위·급탕 시점 설정\n언제든 해당 이벤트에서 제외 가능\n\n단순 타이머 대비 추가 편익으로 평가'],{size:26});
}
// 17
{
let s=add('열 저장량과 이동 가능한 전력량은 다르다','200 L 급탕탱크를 예로 든 물리 계산. 실제 설비 용량 또는 제품 성능 주장이 아님',[],{notes:'물 밀도 약1kg/L, 비열4.186kJ/(kg·K) 가정. 200×4.186×10÷3600=2.3256kWh_thermal. COP3 가정 시 해당 열생산에 드는 전기0.7752kWh. 열손실·혼합·위생운전·출탕 및 COP변동 제외. 전체 탱크를10°C 올릴 수 있는지 여부는 제품별허용온도와 안전설계로 판단.'});
tx(s,'2.33 kWh',70,231,580,105,72,C.teal,true);tx(s,'열량',78,346,490,44,28,C.ink,true);
tx(s,'약 0.78 kWh',670,231,540,105,62,C.teal,true);tx(s,'해당 열을 만드는 전기량',676,346,530,44,28,C.ink,true);
tx(s,'가정: 물 200 L, 온도차 10°C, COP 3',75,434,1100,44,30,C.ink,true);
tx(s,'Q = 질량 × 비열 × 온도차. 전기량은 Q ÷ COP로 환산한다.\n실제 유연성은 열손실·온수 사용·위생 운전·쾌적 제약 때문에 더 작아질 수 있다.',75,498,1090,105,27);
}
// 18
{
let s=add('GT3 경제성: 가격 차이가 좁으면 수익이 사라진다','가상 가구의 연간 민감도. 시장 보상은 0으로 놓고 요금 이동만 계산',[],{foot:'실제 요금·계절·COP·열손실·사용 패턴으로 다시 계산해야 한다. 아래 값은 시장 전망이 아니다.',notes:'동일 열서비스 조건, 고가시간 회피전기1000kWh/년, 고가€0.30/kWh, 예열전기1200kWh/년(20%추가), 연운영료€60, 설치추가비€250 가정. 연총절감=1000*.30-1200*p. p=.05,.15,.25에서240,120,0. 고객순편익=총절감-60 ->180,60,-60. 단순회수기간250/180=1.39년,250/60=4.17년. 할인율·고장·물가·수명은 반영하지 않은 비교용 단순예시.'});
table(s,[['저가 시간 전력단가 가정','€0.05/kWh','€0.15/kWh','€0.25/kWh'],['연간 요금 절감','€240','€120','€0'],['연간 고객 순편익¹','€180','€60','−€60'],['€250 추가 설치비 회수²','약 1.4년','약 4.2년','회수 불가']], [424,242,241,241],{height:292,size:25});
tx(s,'가정: 고가 시간 1,000 kWh 회피, 고가 €0.30/kWh, 예열 소비 20% 증가\n¹ 운영료 €60/년 차감    ² 할인율·수명·고장비 제외한 단순 회수기간',71,543,1120,74,23,C.muted);
}
// 19
{
let s=add('GT2: 인증된 제한 기능 안에서 부하를 배분한다','허브의 가치 가설은 불필요한 공사·차단을 줄이면서 사용자가 원하는 서비스를 유지하는 것',['ul','g100','de'],{foot:'계통 접속 허용, 보호기기 적합성, 현장 시공은 별도 요건이다. 클라우드 최적화만으로 보장할 수 없다.'});
duo(s,['현지 HW 파트너 역할','회로·상·접속점의 승인된 계량\n전력 제한과 고장 시 동작\n접촉기·차단·보호의 적합성\n\n통신 단절, 계량 이상, 재부팅 시험\n\n국가별 인증 범위·설치 조건 확보'],['직접 개발할 배분 기술','하드 제약 안에서 HP·급탕·EV 배분\n\nEV 출발시간, 온도, 급탕 필요량 반영\n\n제어 불가능 부하는 별도 여유 확보\n\n평가: 단순 순차 차단·충전기 관리보다\n불편과 설치비를 함께 줄이는가?']);
}
// 20
{
let s=add('GT4: 계약 충돌을 처리하고 편익을 설명한다','여러 시장의 최대 수익을 더하는 방식으로 사업성을 계산하지 않는다',['stack','tx'],{notes:'NESO의 stacking 문서는 서로 다른 시간의 jumping, 다른 용량의 splitting, 같은 방향 동시 제공의 co-delivery 등을 구분한다. 개별허용 여부를 해당계약·버전별로 저장하는 설계를제안한다. 고객요금절감 추정과 프로그램정산 기준선은 다른 결과이므로 구분한다.'});
table(s,[['공통 기능','구현 범위','시장 파트너 역할'],['제어권 조정','명령 출처·유효기간·우선순위 확인\n동일 자원 중복 예약 방지','고객 동의·접속·계약 조건 제공'],['계약 규칙','시간·자원·방향별 허용 조합 저장\n불명확한 조합은 거래 대상 제외','최신 규칙 해석·참여 자격 확인'],['실행 기록','명령·응답·계측·제외 사유 보관','공식 정산 계량·입찰·청구'],['편익 설명','요금 절감 추정, 수수료, 온도 결과\n불확실성과 누락 데이터 표시','보상 지급·고객 약관']], [240,502,406],{height:380,size:24});
}
// 21
{
let s=add('추가 아이디어 1: 검증된 제어능력 프로파일','기기를 발견한 뒤, 실제로 가능한 제어의 범위를 재사용하는 기술',['gridx','hb','enphase2'],{foot:'신규 개발 가설. 표준 채택·특허성·파트너 구매 의향을 확인한 것은 아니다.'});
lines(s,[['기록할 내용','모델·펌웨어, 읽기/쓰기 권한, 명령 지연, 실측 응답 범위, 시험 조건과 만료 시점'],['사용할 곳','설치 완료 판단, 열 유연성 제공량 산정, 고장 진단, 파트너의 자원 등록 준비'],['검증할 차별성','단순 호환 목록 대비 미응답을 줄이는가? 펌웨어 변경 후 재시험 범위를 줄이는가?']]);
}
// 22
{
let s=add('추가 아이디어 2: 이벤트 후 회복 전력도 예약한다','저가 시간 또는 이벤트 종료 직후 다수 기기가 함께 켜지는 상황을 실증에서 다룬다',[],{foot:'신규 연구 가설. 현장에서 문제가 발생하는 빈도와 크기부터 측정한다.'});
duo(s,['제안 기술','각 가구가 조절한 열량과 회복기한 계산\n\n회복에 쓸 수 있는 전체 전력 예산 설정\n\n온도·급탕이 급한 집부터 순차 재개\n\n기기 휴지시간을 지키며 재기동 분산\n클라우드 단절 시 로컬 기본 동작'],['사업화 판단','구매자 후보: 소매사·집합사업자\n\n비교군: 동시 복귀, 고정 무작위 지연\n\n평가: 최대 회복 kW, 온도 이탈, 비용\n\n차이가 작으면 독립 과제로 만들지 않고\nGT3의 운전 기능으로 포함']);
}
// 23
{
let s=add('안전 제한과 경제 최적화의 책임을 분리한다','가격·AI·원격 명령은 로컬 보호 기능을 우회하지 못하도록 설계',['ul','eebus2','cra']);
table(s,[['위치·책임','동작','기술 선택'],['기기·인증 보호계층','과전류, 기기 보호, 승인된 전력 한계','인증 하드웨어·결정론적 로직'],['로컬 게이트웨이','명령 권한·유효기간 검증\n통신 실패 시 기본 운전 복귀','서명 정책, 상태기계, 이벤트 로그'],['로컬·클라우드 최적화','온도·급탕·배터리·EV 제약 안의 배분','제약 최적화, 열상태 추정'],['사업 운영계층','계약·가격 갱신, 편익 설명, 고객 지원','버전 관리, 정산 파트너 연계']], [312,483,353],{height:382,size:24});
}
// 24
{
let s=add('직접 개발은 현장 데이터가 쌓이는 부분에 집중','범용 연결과 시장 접속까지 모두 직접 만들면 7명 조직의 범위가 과도해진다',[]);
table(s,[['선택','범위','결과물'],['직접 개발','시험 계획·응답 판정\n열 유연성·회복 제어','현장 시험 엔진\n가용량 산정 및 제어 모듈'],['표준·구매 활용','통신 스택, 기본 게이트웨이\n계량·보호 부품, 기상·요금 피드','조합별 호환·버전 관리'],['제휴','인증 전력 허브, 설치 채널\n시장 등록·입찰·정산','공동 시험·제품 공급·운영 계약'],['초기 보류','범용 V2H, 신규 범용 배전반\n모든 브랜드·국가의 동시 지원','선행 조건을 만족할 때 재검토']], [258,506,384],{height:375,size:25});
}
// 25
{
let s=add('실증은 기술 성능과 구매 의향을 함께 확인한다','아래 규모는 초기 계획안이다. 통계 검정력과 계절 범위를 확인한 뒤 확정',[],{foot:'독일 현장 30건과 영국 40가구는 초기 검증 규모 제안이며 시장 성능을 일반화하는 표본이 아니다.'});
table(s,[['트랙','초기 범위 제안','반드시 확보할 비교'],['GT1 / 독일','설치업체 2곳, 현장 30건\n우선 기기 조합 3개','동일 작업 범위의 기존 절차\n총 시간·재방문·사용료 의향'],['GT3 / 영국','40가구, 난방기 12주 이상\n계량·온도·HP 상태 동의 확보','기존 타이머와 대조 운전\n요금·쾌적·회복 부하'],['GT2 / 파트너 시험실','인증 파트너 1곳\n계량·통신·제어 고장 주입','기존 부하관리·증설 견적\n현지 승인 경로와 비용'],['GT4 / 계약 검증','시장 파트너 1곳\n기존 기록으로 정산 재현','공식 기준선과 재계산 차이\n중복 참여 규칙·고객 설명']], [262,443,443],{height:383,size:24});
}
// 26
{
let s=add('확대 기준은 과제 착수 때 수치로 합의한다','초기 목표 제안. 사전 기준선·측정 오차·현장 조건을 확인한 후 확정',[],{notes:'성능 목표는 보장이나 현실증 결과가 아니다. 제한된 시험의0건실패는 실제모집단에서위험0을뜻하지않는다. 안전관련실패가나면확대중단후원인교정·재검증. 온도이탈비율분모는재실및제어대상시간, 허용온도폭은가구별사전합의. 고객순편익95%구간은가구군수준추정이며개별가구확정액아님.'});
table(s,[['과제','확대 판단 기준 제안','보류·축소 조건'],['GT1','설정·시험 중앙시간 40% 이상 단축\n재방문 증가 없음, 유료 도입 의향','지원 범위 확대 비용이 절감액 초과'],['GT3','단순 타이머 대비 순편익 추정\n하한도 양수, 온도 이탈 악화 없음','추가 편익이 운영·지원비보다 작음'],['GT2','고장시험 통과, 현지 승인 경로 확보\n설치 총비용이 비교 대안보다 유리','인증 일정·원가로 경쟁 불가'],['GT4','중복 지급·명령 충돌을 검출\n파트너 정산 결과 차이를 설명','데이터·계약 접근을 확보하지 못함']], [228,550,370],{height:380,size:24});
}
// 27
{
let s=add('2027–2030 로드맵은 증거가 쌓일 때 확장한다','연도는 개발·상용화 목표안이며 확정 출시일 또는 매출 계획이 아님',[]);
table(s,[['기간','주요 결과물','다음 단계 조건'],['2026 Q4','현장 시간·요금·기기 권한 조사\n실증 파트너와 비교 기준 합의','데이터 접근·시험 허용·구매자 확인'],['2027','GT1 지원 조합 3개 시제품\nGT3 난방기 실증·가격 검증','두 과제 중 최소 1개 유료 도입 근거'],['2028','선별 제품화, 네덜란드 지역화\nGT2 파트너 HW 통합·인증 시험','운영비 포함 이익·채널 재판매 확인'],['2029–2030','검증된 OEM·국가로 확장\nGT4 추가 계약·시장 어댑터','기존 고객 유지와 설치·지원 품질']], [221,548,379],{height:382,size:24});
}
// 28
{
let s=add('7명 기준의 역할과 사업 모델','내부 인력·기술 보유 현황은 미확인. 7명 체제를 가정한 배치안',[]);
duo(s,['개발 인력 배치  총 7명','시스템·안전 아키텍처  1명\n엣지 통신·설치 시험  2명\n열 제어·최적화  2명\n데이터·편익 검증  1명\n현장 실증·시험 자동화  1명\n\n영업·인증·보안 검토는 별도 지원 필요\n기존 과제 병행 시간은 제외한 가정'],['판매 단위와 지불 이유','설치업체 / 건별 사용료\n작업시간·재방문 감소로 회수\n\nOEM / 탑재료와 유지관리료\n제품·펌웨어별 검증 비용 절감\n\n소매사 / 활성 기기별 운영료\n이행 가능한 유연성·운영 품질에 지불\n성과배분은 정산·책임 합의 후 적용'],{size:25});
}
// 29
{
let s=add('다음 90일 조사에서 투자 여부를 결정한다','공개자료로 확인하지 못한 현장 비용·권한·가격을 확보하는 단계',[]);
table(s,[['기간','확인할 미지수','남길 증거'],['1–30일','설치 지연의 실제 원인과 재방문 비용\n고객·기기 데이터 접근 권한','설치 동행 관찰 10건\n작업 분류표·익명 비용 기록'],['31–60일','기기 응답 시험과 열 제어 구현 가능성\n현행 제품 대비 차이','우선 조합 3개 시험\n기존 방식과 결과 비교'],['61–90일','구매자가 지급 가능한 가격\n파트너 인증·운영·지원 비용','유료 실증 제안과 가격 면담\n부품·인증·지원 원가 견적'],['90일 판단','추가 편익이 대안 대비 충분한가?','과제별 계속 / 범위 축소 / 중단']], [215,528,405],{height:380,size:24});
}
// 30
{
let s=add('이번에 결정할 투자 범위','독일·영국의 두 실증으로 기술과 지불 의향을 함께 확인',[],{dark:true});
lines(s,[['착수  GT1과 GT3의 공통 기반','기기 응답 검증, 제어능력 프로파일, 열 유연성·회복 운전'],['조건부 추진  GT2·GT4','인증 HW와 현지 시장 파트너가 확보한 범위 안에서 연계'],['확대 조건','단순 대안보다 높은 고객 순편익, 유료 도입 근거, 반복 가능한 설치·지원 비용']],{dark:true,gap:126});
}
// 31 appendix
{
let s=add('부록  제도별 제한 대상이 서로 다르다','하나의 “국가별 전력 제한값”으로 추상화하면 잘못 제어할 수 있다',['de','g100','ul','stack','tx']);
table(s,[['제도·표준','구분해야 하는 대상','구현 시 확인할 것'],['독일 §14a','대상 제어가능 설비의 계통 수전\n일반 가전·PV 역송과 구분','설비 조합·설치 시점·DSO 방식'],['영국 G100','접속점의 수전·역송 제한 체계','해당 판·접속 조건·고장 동작'],['미국 PCS','해당 인증 범위의 전력·회로 제한','현지 법규 채택·유틸리티·현장 승인'],['영국 DFS / 텍사스 ADER','서비스·계약별 자원 참여·정산','동의·계량·기준선·중복 참여']], [274,498,376],{height:377,size:24});
}
// 32
{
let s=add('부록  표준 지원과 실제 제어 가능성의 차이','기기·펌웨어·권한·지역 조건을 조합별로 검증하는 능력이 필요',['matter','eebus','eebus2','cra']);
table(s,[['확인 사항','의미','R&D에 반영할 일'],['Matter 1.5 에너지 확장','요금·계량·계통 정보 교환 기능\n2025.11 발표','표준 기능 존재와 설치기기 지원을\n별도 호환 시험으로 확인'],['EEBUS 전력 제한','LPC·LPP 등 용도별 제한 기능','모든 기기에서 동일 기능·응답을\n보장한다고 가정하지 않음'],['EU Cyber Resilience Act','주요 의무 2027.12.11 적용\n보고 의무 2026.09.11부터','보안 업데이트·취약점 대응·\n지원기간 비용을 제품 설계에 포함']], [316,416,416],{height:339,size:24});
tx(s,'출하 이후에도 펌웨어·설정·시험 기록을 관리해야 설치 자동화의 가치가 유지된다.',70,588,1100,31,24,C.teal,true);
}
// 33
{
let s=add('부록  편익 검증과 시장 정산을 구분한다','같은 집에서 제어하지 않았을 때의 전기요금은 직접 관측할 수 없다',['heat','heat2','stack'],{notes:'실증제안: 기상·재실·기기구성에따른층화, 사전분산측정후표본수결정, 사전등록한기간·제외기준으로대조. 교차설계는열관성의잔류효과와washout처리. 탈락·통신결측은제거만하지않고발생률·민감도보고. 사용요금절감은예열과회복비용·추가전력·수수료 포함. 계량경계·시간대·DST·세금·기본료처리 명시.'});
duo(s,['고객 순편익 검증','날씨·재실·요금이 비슷한 대조 운전\n또는 무작위 배정·교차 설계 사용\n\n예열과 회복 기간까지 전체 비용 포함\n열관성의 잔류 효과를 분리\n\n평균뿐 아니라 분산·불확실성 제시\n온도 이탈·거부·결측을 함께 보고'],['프로그램 정산','사업자가 정한 공식 기준선 적용\n\n인정 계량기·구간·이벤트 조건 확인\n\n고객 요금 절감과 시장 보상 별도 표시\n동일 자원·시간의 이중 집계 차단\n\n추정 절감액을 확정 정산액처럼\n표시하지 않음'],{size:26});
}
// 34
{
let s=add('부록  특허와 차별성은 시험 방법까지 좁혀 검토한다','선행기술 샘플 확인이며 특허성·권리범위·실시 자유의 종합 판단이 아님',['patent','eebus2','heat']);
table(s,[['영역','이미 존재하는 선행기술','새로 검토할 좁은 가설'],['CT 극성·연결 자동 확인','US10191102B2\nCT 극성 자동 보정','가정 내 여러 기기의 제한된 자극으로\n제어권·계측점·응답을 함께 판정'],['열 유연성 제공','HeatFlex 현장 실증\n기존 모델·제어 연구','설치 시험 결과를 이용한\n가용량 불확실성의 초기 축소'],['제약·통신 실패 대응','EEBUS LPC 등 공개 사양','펌웨어 변화 후 위험에 따라\n재시험 범위를 선택하는 방법']], [260,418,470],{height:354,size:24});
tx(s,'“자동화·AI·멀티브랜드”라는 표현만으로 2–3년의 기술 우위나 특허를 주장하기 어렵다.',70,586,1110,34,23,C.teal,true);
}
// 35
{
let s=add('부록  기존 12개 아이디어의 처리 방향','제공 HTML의 후보를 공통 과제·지역 확장·보류 대상으로 재배치',['ref']);
table(s,[['기존 후보','이번 반영','재검토 조건'],['T1 자동 시운전 · T2 규칙 코드화','GT1 핵심 + 국가별 설정 모듈','지원 조합의 현장 시간 절감'],['T3 열 유연성 · T11 성능 관리','GT3 핵심 + 응답·원격 진단','타이머 대비 순편익·지원비'],['T4 전력 제한 · T5 증설 회피','GT2 인증 HW 제휴','실제 공사 견적·승인 경로'],['T7 제어권 · T8 편익 검증','GT4 공통 운영 기능','파트너 계약·데이터 접근'],['T6 백업 · T9 레거시\nT10 V2H · T12 임대형 소형팩','지역 옵션 또는 초기 보류','설치권·보증·인증·추가 HW\n구매자와 원가 확인']], [455,319,374],{height:385,size:23});
}
// 36
{
let s=add('부록  추가 개발 전 확인할 전제','주거형태와 기기 보호 조건이 충족돼야 상품을 판매할 수 있다',['ref','ul','enphase2']);
table(s,[['후보','확인할 전제','초기 판단'],['백업·정전 대응','정전 빈도·중요 부하·잔량 선호\n전환 장치·섬운전 가능 제품','미국 파트너 상품의 옵션'],['레거시·IR 어댑터','명령 성공을 확인할 상태 피드백\n설정 충돌·OEM 보증','지원 모델을 제한한 설치 옵션'],['V2H','차종·충전기·전환장치 조합\n보증·접속·계약 승인','공통 기능 선행 개발은 보류'],['임대형·플러그인 팩','설치권·계량 경계·역송 제한\n현지 접속 요건·회수 비용','대상 국가와 인증 경로 확정 후']], [260,528,360],{height:374,size:25});
}
// 37–39 source register: titles on slide, complete clickable URLs in notes.
const groups=[
 ['조사 근거 1  국가 제도·요금',['de','uk','stack','fr','nl','it','it2','es','ca','tx']],
 ['조사 근거 2  지역 경제성과 경쟁 제품',['irs','ae','sa','qa','gridx','hb','kraken','enphase','enphase2','span']],
 ['조사 근거 3  인증·실증·표준·선행기술',['cd','ul','g100','heat','heat2','eebus','eebus2','matter','cra','patent']]
];
for(let [title,keys] of groups){let s=add(title,'각 번호의 원문 URL은 해당 슬라이드 발표자 노트에 수록. 조사 기준일 2026.09.30',keys);keys.forEach((k,i)=>{let y=218+i*38;tx(s,src[k][0],69,y,72,32,20,C.teal,true);tx(s,src[k][1]+'  '+src[k][2],151,y,1050,32,21,C.ink);});}

await fs.mkdir(B+'/previews',{recursive:true});
await fs.mkdir(W+'/output',{recursive:true});
await (await PresentationFile.exportPptx(P)).save(B+'/candidate.pptx');
console.log('EXPORTED '+pages.length+' slides');
await fs.writeFile(B+'/sources.json',JSON.stringify(src,null,2));
await fs.writeFile(B+'/manifest.json',JSON.stringify({titles:pages.map(x=>x.title),tableOwners,output:OUT},null,2));
for(let i=0;i<pages.length;i++){
 const png=await P.export({slide:pages[i].s,format:'png',scale:1});
 await fs.writeFile(B+'/previews/'+String(i+1).padStart(2,'0')+'.png',new Uint8Array(await png.arrayBuffer()));
 const lay=await pages[i].s.export({format:'layout'});await fs.writeFile(B+'/previews/'+String(i+1).padStart(2,'0')+'.json',await lay.text());
 console.log('RENDER '+(i+1));
}
await finalizePresentation({workspaceDir:W,candidatePath:B+'/candidate.pptx',finalPath:OUT,pythonExecutable:PY,integrityValidatorPath:SK+'/container_tools/inspect_presentation_package_integrity.py',layoutValidatorPath:SK+'/container_tools/inspect_presentation_layout_geometry.py',layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:tableOwners,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:B+'/validation-final.json'});
console.log('FINAL '+OUT);
