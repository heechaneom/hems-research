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

// 01
{
let s=slide('Energy Pack의 다음 성장 기회','', ['meeting'],true);
txt(s,'국가별 Pain Point에서 찾는\n2027–2028 신기술 과제',66,227,1120,170,57,C.white,true);
txt(s,'설치 부담을 줄이고, 냉난방까지 제어하고,\n고객에게 남는 이익을 증명하는 사업',69,447,1100,100,30,'#BFD8DC');
txt(s,'CTO 의사결정용 제안 초안',69,597,1020,40,24,C.gold);
}
// 02
{
let s=slide('추천: 독일·영국에서 2개 과제를 우선 검증','네덜란드를 후속 시장으로 준비하고, 7명은 공통 기술 개발에 집중',['meeting','de','uk','nl']);
blocks(s,[['과제 A  설치 자동화 모듈','혼합 제조사 기기의 연결·설정·검증을 자동화해 설치업체의 작업시간과 재방문을 줄인다.'],['과제 B  냉난방 유연성 제어','쾌적함을 지키면서 히트펌프·급탕·ESS의 운전을 조정하고 실제 순편익을 입증한다.'],['조건부 확장  정전 대응·전력시장 참여','공통 기술을 확보한 뒤 미국·호주 파트너와 확장한다. 초기부터 모든 국가를 개발하지 않는다.']],{gap:132});
}
// 03
{
let s=slide('사업과 기술의 책임 범위','스마트 코티지는 적용 상품, Energy Pack은 판매 패키지, HEMS는 공통 제어 기술',['meeting']);
table(s,[['구분','고객에게 파는 것','이번 과제의 기여'],['스마트 코티지','주거 공간과 설비를 결합한 상품','에너지·스마트홈 통합 실증'],['Energy Pack','PV·ESS·히트펌프 등 구성 상품','설치·운영 모듈로 판매 경쟁력 강화'],['HEMS 공통 기술','기기 연결, 에너지 운전, 운영 서비스','국가를 바꾸어도 재사용하는 기술']], [230,445,477],{height:337,size:25});
note(s,'미국 스마트 코티지 판매는 회의에서 제외한 범위. 미국 기회는 Energy Pack·기술 파트너십으로 별도 검토.');
}
// 04
{
let s=slide('국가별 기회: 유럽 4개국','공식 제도·소비자 조사에 근거한 문제와 사업 가설을 구분',['de','de14','uk','fr','nl']);
table(s,[['국가·대상','확인한 변화','지불 이유에 대한 가설'],['독일 / HP·EV 보유 주택','동적요금 제공 의무, §14a 제어','설치업체가 연동·검증 비용 지불'],['영국 / 전기난방·EV 가구','TOU 미가입자 중 전환 의향 17%','요금사업자가 절감 신뢰·운영을 구매'],['프랑스 / 급탕·전기난방 가구','2025–2027 비첨두 시간대 변경','기존 기기 설정 변경을 원격으로 처리'],['네덜란드 / PV 보유 가구','2027.1.1 상계제도 종료','자가소비 확대를 위한 열부하 제어']], [300,430,422],{height:390,size:23});
note(s,'TOU: 시간대별 요금. 제도 변화가 곧 구매수요를 의미하지는 않으며, 지불 의향은 현지 검증 대상.');
}
// 05
{
let s=slide('국가별 기회: 유럽 외 5개국','전국 평균보다 고객의 주거형태·설비·요금 계약이 중요',['au','ca','uae','kr','za']);
table(s,[['국가·범위','확인한 배경','검증할 상품'],['호주 / PV·배터리 보유 가구','2025.7부터 가정용 배터리 지원','배터리·냉난방 통합 운전'],['미국 / 캘리포니아 우선','NBT로 태양광 수출 보상 구조 변화','열부하·배터리 운전과 백업 잔량 관리'],['UAE / 개별 냉방 빌라','DEWA가 냉방 설정·운영 효율 강조','온습도 기반 쾌적 냉방 제어'],['한국 / 공동주택 사업자','2024년 전체 주택 중 아파트 65.3%','공용 설비·충전기 묶음 운영'],['남아공 / 지역 한정','2026.9.25 기준 497일 부하차단 없음','정전 상품은 지역 장애 수요부터 재확인']], [300,430,422],{height:403,size:22});
note(s,'남아공의 전국 순환 부하차단(loadshedding)과 지역 공급제한(load reduction)은 다른 현상.');
}
// 06
{
let s=slide('독일: 설치업체의 복잡성을 줄이는 제품','동적요금·히트펌프·EV·계통제어를 한 현장에서 다뤄야 하는 고객',['de','de14']);
cols(s,[['고객과 기존 대안','지불자 후보: Energy Pack 설치업체·OEM\n\n기존 대안: 기기별 앱과 매뉴얼,\n이미 판매 중인 HEMS\n\n문제 가설: 연결은 가능해도\n배선·제어 권한·설정 검증이 번거롭다.'],['제안: 설치 자동화 모듈','지원 기기를 발견하고 제어 능력을 확인\n\n승인된 지역 규칙으로 설정을 생성\n\n통신·계측·계통 명령 시험 후\n설치 완료 보고서를 자동 발급\n\n1차 목표: 설정·검증 시간 50% 단축']]);
note(s,'§14a는 모든 가전 대상이 아님. 2024년 이후 신규 설치 등 적용 조건을 확인하며, 목표치는 제안값.');
}
// 07
{
let s=slide('영국: 절약 가능성과 실제 가입 사이의 간극','Ofgem 2025 조사, 2026년 발표. 아래 두 비율은 모집단이 같지 않음',['uk','oct']);
txt(s,'49%',69,226,460,106,82,C.teal,true);txt(s,'TOU가 비용을 줄일 수 있다고 보는 소비자',72,348,495,95,27);
txt(s,'17%',665,226,465,106,82,C.teal,true);txt(s,'TOU 미가입자 중\n향후 전환 의향이 있는 소비자',667,348,490,95,27);
txt(s,'제안: 절감 증명 + 쾌적함을 지키는 자동운전',70,487,1110,51,31,C.ink,true);
txt(s,'전력 소매사와 판매하고, 고객에게는 운전 전후 순편익·수수료·온도 이탈을 함께 공개',72,557,1110,65,26);
}
// 08
{
let s=slide('프랑스·네덜란드: 제도 변화가 운전을 바꾼다','요금 데이터 갱신만으로 끝내지 않고 급탕·냉난방 운전까지 연결',['fr','nl']);
cols(s,[['프랑스  스케줄 변경 자동화','확인: 1,100만 고객의 비첨두 시간대가\n2025.11–2027.10 순차 변경\n\n대상: 고정 타이머로 급탕·난방하는 집\n\n제안: 공식 계약 정보를 확인한 뒤\n기기 스케줄을 변경하고 결과 검증\n\n요금 변경 대응은 지역화 기능으로 개발'],['네덜란드  태양광을 열로 저장','확인: 2027.1.1 상계제도 종료 예정\n\n대상: PV와 급탕탱크·HP 보유 가구\n\n제안: 잉여 태양광으로 급탕·예열하고\n필요한 경우 배터리와 함께 최적화\n\n핵심 연구: 집별로 쓸 수 있는 열 유연성']]);
note(s,'자가소비 확대의 경제성은 수출 보상·수수료·열손실에 따라 달라짐. 배터리 구매를 전제하지 않음.');
}
// 09
{
let s=slide('미국·호주: 배터리의 돈벌이와 백업을 함께 관리','설치된 배터리를 활용하되 현지 전력사업자를 통해 진입',['ca','au','vpp']);
cols(s,[['미국  캘리포니아부터','확인: NBT는 2023.4.15 이후\n연계 신청 고객에 적용하는 후속 제도\n\n문제 가설: 수출·자가소비·백업 사이\n운전 선택을 고객이 이해하기 어렵다.\n\n제안: 배터리 잔량과 냉난방 부하를\n함께 조절하는 백업 운전 모듈'],['호주  추가 배터리 설치 고객','확인: 가정용 배터리 지원 제도 시행\n\n문제 가설: 기존 PV·신규 배터리·\n냉난방을 따로 운전한다.\n\n제안: 설치 자동화와 통합 운전 제공\n\nVPP 수익은 계약 확보 후에만 반영']]);
note(s,'VPP: 여러 집의 유연한 전력자원을 묶는 가상발전소. 국가 전체가 단일 시장인 것으로 가정하지 않음.');
}
// 10
{
let s=slide('UAE·한국: 비용 부담자와 제어권자가 중요','가정용 구독에 고정하지 않고 건물 운영·설비 판매 채널을 검토',['uae','kr']);
cols(s,[['UAE  개별 냉방 빌라','대상: 냉방비를 직접 부담하는 소유자\n또는 다수 주택 운영사\n\n제안: 습도·재실을 반영한 냉방과\n설비 이상 감지를 묶어 판매\n\n주의: 지역냉방 공동주택에는\n같은 절감·제어 구조를 적용하기 어렵다.'],['한국  공동주택 운영사','대상: 건설사·관리사업자·임대 운영사\n\n제안: 공용부 부하·EV 충전·\n적용 가능한 냉난방 설비를 묶어 운영\n\n입주민 개인 구독보다\n설비 도입·운영 계약에서 비용 회수\n\n계량·요금·의사결정 권한을 먼저 확인']]);
}
// 11
{
let s=slide('경쟁 제품이 이미 제공하는 기능','자동운전·동적요금·전력 허브라는 이름만으로 신기술을 주장하기 어렵다',['sma','tesla','oct','sch']);
table(s,[['경쟁 대안','공식 제품에서 확인한 기능','우리의 추가 검증 과제'],['SMA Home Manager','PV·배터리·HP·EV 및 요금 연동','혼합 기기 설치시간·재방문 감소'],['Tesla Powerwall','시간대별 가격 기반 충·방전','냉난방 포함 추가 순편익'],['Octopus Heat Pump','요금·전력망을 반영한 난방 운전','타 기기 조합에서의 성능·채널'],['Schneider Home','패널·태양광·배터리·충전 통합','기존주택 개조 비용·설치 범위']], [280,440,432],{height:365,size:24});
note(s,'제품별 국가·모델·연동 조건이 다름. 경쟁사에 해당 기능이 없다고 단정하지 않고 동일 조건 실측으로 비교.');
}
// 12
{
let s=slide('발굴한 신기술 후보 6개','신규성은 연구 가설이며, 선행 특허·제품 비교를 통해 확인');
table(s,[['후보','해결할 문제','우선 판단'],['A 설치 자동화 모듈','연결 후 설정·시험·재방문 비용','우선 개발'],['B 열 유연성 제어 엔진','쾌적함 손실 없이 부하 이동','우선 개발'],['C 백업 지속시간 관리','ESS 용량이 있어도 필요한 생활 유지 불확실','파트너와 후속 실증'],['D 절감 증명·정산 엔진','효과를 믿기 어려워 유료 전환 지연','A·B의 필수 운영 기능'],['E 교체형 전력 허브','기존주택 개조·배선·설치 비용','인증 HW 파트너와 검토'],['F 지역 규칙 자동 검증','국가별 요금·제어 규칙 오류','지역화 기반으로 개발']], [310,580,262],{height:416,size:23});
}
// 13
{
let s=slide('A. 설치 자동화: 연결 후 검증까지 제품화','판매 단위: 게이트웨이·계측 인터페이스·설치 SW 라이선스');
blocks(s,[['새로 연구할 것','기기 설명 정보를 공통 형식으로 변환하고, 작은 시험 신호로 제어 방향·응답·계측 관계를 확인'],['HW와 SW의 결합','통신 어댑터·계측 입력을 가진 모듈에 기기 발견, 설정 생성, 오류 진단, 시험 이력을 탑재'],['검증 목표와 범위','지원 기기 10종의 설정·검증 시간 50% 단축, 지원 조합 자동 설정 성공률 95% 이상']],{gap:130});
note(s,'목표치 제안. 전기 배선·안전검사는 전문가가 수행하며, “10분 설치”는 전체 전기공사 시간으로 주장하지 않음.');
}
// 14
{
let s=slide('B. 열 유연성: 집이 제공할 수 있는 전력 조절량','판매 단위: 히트펌프·급탕·ESS 제어 기능과 전력사업자 연계 서비스');
blocks(s,[['새로 연구할 것','실측 온도·습도·기기 상태로 “얼마나 오래, 얼마나 줄여도 괜찮은지”를 지속 추정'],['제어 방법','열 상태 추정 + 불확실성을 반영한 예측제어 + 제품 운전 한계. 온라인 탐색은 안전 범위 안에서만 수행'],['검증 목표','재실 중 고객 합의 온도 범위 유지율 95% 이상, 기존 운전 대비 피크 구간 평균 전력 20% 감소']],{gap:130});
note(s,'목표치 제안. 요금 절감과 에너지 절감은 다름. 열손실·추가 소비를 포함한 순편익이 양수여야 출시.');
}
// 15
{
let s=slide('C·E. 백업 제어와 전력 허브는 선택적으로 확장','배터리 셀 신규 개발보다 기존 인증 전력장치와 가전 제어의 결합을 먼저 검증');
cols(s,[['C  백업 지속시간 관리','고객이 보호할 냉장·조명·냉방을 선택\n\n배터리 잔량·온도·부하별 운전 제약으로\n지속시간과 불확실성 범위를 제시\n\n목표: 동일 배터리로 중요부하\n유지시간 20% 증가\n\n정전 예측 자체를 판매 가치로 삼지 않음'],['E  교체형 전력 허브','기존 분전 구조에 적용 가능한\n계측·제어 모듈을 인증사와 공동 설계\n\n차단·계통 분리 기능은 인증 장치가 담당\n\n목표: 배선·설치 공수 30% 단축\n\n신규 배터리·인버터 전체 개발은\n7명 과제의 초기 범위에서 제외']]);
note(s,'모든 수치는 실증 전 목표. 백업 전환 성능·부하 우선순위는 인증 HW 구성과 고객 계약에 종속.');
}
// 16
{
let s=slide('공통 기술과 국가별 변경 영역','AI 역할을 한정하고, 전력기기 실행은 검증된 제어 규칙으로 관리');
table(s,[['공통 개발 영역','국가별 변경 영역','실행 원칙'],['기기 능력·상태의 공통 표현','지원 제조사·프로토콜·모델 목록','목록 밖 기기는 읽기만 또는 수동 검토'],['열 상태 추정·안전 제약 제어','기후·건물·고객 쾌적 조건','제품 제어기의 보호 기능 우선'],['설정 생성·자동 시험','요금·계통 신호·인증 시험','승인한 버전만 배포·되돌리기 지원'],['순편익 측정·운영 기록','계약·정산 방식·개인정보 조건','기준선과 수수료를 함께 공개']], [405,390,357],{height:370,size:23});
note(s,'언어모델은 문서 해석·설정 후보·진단 설명에 활용. 요금·규제를 임의 해석해 전력기기에 직접 명령하지 않음.');
}
// 17
{
let s=slide('우리의 강점은 제품 제어 깊이로 입증해야 한다','회의에 언급된 실증 경험은 활용하되 자료 확인 전 성과 수치를 확정하지 않음',['meeting']);
cols(s,[['활용할 수 있는 자산 후보','가전·히트펌프의 실제 제어 경험\n\n제품 운전 제약과 내부 상태 데이터\n\n설치·서비스 채널과 현장 데이터\n\n자사 기기 깊은 제어 + 타사 기기 표준 연동\n조합으로 차별화 가설 수립'],['CTO 보고 전 확보할 증거','개발 완료 기술·적용 모델·완료 연도\n\n실증 대상·기간·대조군·원시 로그\n\n제품 내부 상태 접근 권한과 API 범위\n\n자사 기기 / 표준 연동 / 경쟁 제품의\n동일 조건 성능 차이']]);
note(s,'“2~3년 기술 격차”는 현시점의 입증된 사실이 아님. 설치·제어 데이터와 시험 자산 축적을 목표로 설정.');
}
// 18
{
let s=slide('시장 선택: 개발은 독일·영국, 후속은 네덜란드','공식 자료와 사업 구조를 바탕으로 한 정성 판단. 확정 시장 순위가 아님');
table(s,[['판단 기준','독일','영국','네덜란드','미국·호주'],['구체적인 진입 계기','계통·요금 통합','가입·절감 신뢰','2027 상계 종료','배터리 운영'],['우리 제품과의 연결','높음 / HP·설치','높음 / HP·운영','조건부 / 열부하','조건부 / 제휴'],['우선 확보할 채널','설치업체·OEM','전력 소매사','PV·HP 설치업체','VPP·인버터 업체'],['초기 의사결정','1순위 현장 검증','1순위 현장 검증','공통 기술로 확장','계약 후 후속']], [264,222,222,222,222],{height:359,size:23});
note(s,'채널 계약과 기기 접근 권한을 확보하지 못한 국가는 뒤로 미룸. 국가 수보다 유료 전환 근거를 우선.');
}
// 19
{
let s=slide('수익 구조: 설치 모듈 판매와 운영 계약','가정용 HEMS 단독 구독 외에도 지불자를 선택할 수 있다');
table(s,[['상품','지불자 후보','왜 돈을 내는가'],['설치 자동화 모듈','설치업체·제품 OEM','작업시간·재방문 비용 감소'],['열 유연성 제어 기능','HP·Energy Pack 사업부','제품 판매·가격 경쟁력 개선'],['운영·효과 측정 서비스','전력 소매사·주택 운영사','고객 유지·운영 비용·절감 신뢰'],['VPP 연계 기능','현지 집합사업자','검증 가능한 유연성 자원 확보']], [340,350,462],{height:365,size:24});
note(s,'초기 매출에는 VPP 정산 수익·추가 가전 판매 효과를 제외. 실증 계약에서 지급 조건부터 검증.');
}
// 20
{
let s=slide('2028년 매출 시나리오: 기준안 31.8억원','모든 물량·단가는 작성자 가정. 시장규모 추정이나 회사 사업계획이 아님');
table(s,[['지역','신규 설치','모듈 매출','운영 매출','합계'],['독일','4,000호','12.00억원','0.72억원','12.72억원'],['영국','3,000호','9.00억원','0.54억원','9.54억원'],['네덜란드','3,000호','9.00억원','0.54억원','9.54억원'],['합계','10,000호','30.00억원','1.80억원','31.80억원']], [265,202,225,230,230],{height:350,size:24});
note(s,'모듈 30만원/호, 운영 3.6만원/호·년. 연중 균등 설치로 평균 과금 0.5년. Energy Pack 본체 매출 제외.',605);
}
// 21
{
let s=slide('기준안은 연구개발비까지 회수하지 못한다','물량 확대·원가 절감·기존 제품 사업 기여 중 최소 하나를 입증해야 지속 가능');
table(s,[['2028년 가정','보수안','기준안','확장안'],['신규 설치','3,000호','10,000호','30,000호'],['증분 매출','9.54억원','31.80억원','95.40억원'],['변동비 차감 후 기여이익','3.06억원','10.20억원','30.60억원'],['가정 고정비 12.5억원 차감','−9.44억원','−2.30억원','18.10억원']], [440,238,237,237],{height:327,size:24});
txt(s,'같은 가정의 손익분기: 연간 약 1.23만 호',68,571,1135,46,30,C.teal,true);
note(s,'모듈 변동비 21만원, 운영 변동비 연 1.2만원/호. 고정비: 7명×1.5억원+외부 실증·인증 2억원.',624);
}
// 22
{
let s=slide('12주 검증: 사업성과 기술성을 함께 판정','표본은 탐색 실증 규모이며, 출시 성능 검증은 2027년 계절 실증에서 확대');
table(s,[['기간','수행 내용','계속할 조건'],['1–3주','국가별 고객 10명·설치자 5명 인터뷰\n기기 접근·데이터·요금 계약 확인','유료 실증 의향 파트너 2곳'],['4–6주','지원 10기종 설치 시험\n기존 HEMS와 설정·검증 시간 비교','설정 시간 50% 감소 가능성'],['7–10주','독일·영국 각 10호 탐색 실증\n교차 운전·고객 수동개입 기록','순편익 양수, 쾌적 조건 유지'],['11–12주','가격·설치 공수·문의 비용 확인\n계절 확대 실증과 제품화 판단','유료 계약 또는 명확한 구매 조건']], [180,566,406],{height:386,size:23});
note(s,'접근 권한 미확보·고객 순편익 음수·반복 안전 제약 위반이면 해당 기능 또는 시장 범위를 축소.');
}
// 23
{
let s=slide('2026년 말부터 2028년까지의 개발 로드맵','기존 기술 적용과 새 연구 과제를 분리하고, 연차별 중단 기준을 둔다');
table(s,[['시기','제품·연구 결과','사업 검증'],['2026년 4분기','기존 기술 목록·원시 실증 확인\n국가별 파트너·유료 실증 설계','설치 공수·구매자 문제 확인'],['2027년 상반기','설치 모듈 시제품, 10기종 지원\n열 유연성 제어 초기 모델','12주 탐색 실증과 유료 계약'],['2027년 하반기','총 100호 목표 계절 실증\n인증 설계·경쟁 제품 비교','반복 설치 원가·순편익 검증'],['2028년','검증 통과 시 제품화·국가 확장\n백업·VPP 기능은 파트너와 추가','연 1만 호 기준안을 계약으로 검증']], [232,575,345],{height:385,size:23});
note(s,'100호·1만 호는 계획 가정. 현장 모집과 인증 일정에 따라 단계 전환. 완료 연도는 내부 자료로 확정.');
}
// 24
{
let s=slide('7명 투입안: 두 과제에 공통 인력을 배치','전력장치 전체 개발·국가별 영업·인증을 연구인력 7명이 모두 맡지 않는다');
table(s,[['역할','인원','책임'],['PL·제품/파트너 요구','1명','범위·사업 지표·파트너 계약 조건'],['기기 연동·설치 자동화','2명','프로토콜, 설정·시험, 오류 진단'],['열 상태 추정·제어','2명','제품 운전 제약, 유연성 추정, 순편익'],['임베디드·HW 인터페이스','1명','게이트웨이·계측·인증사 기술 연계'],['실증·평가·데이터 품질','1명','대조 실험, 현장 로그, 비교 검증']], [445,140,567],{height:385,size:23});
note(s,'사업부: 영업·기기 접근 / 현지 파트너: 설치·정산 / 인증사: 전기 안전·전파. 역할·비용 합의가 착수 조건.');
}
// 25
{
let s=slide('이번 보고에서 결정할 세 가지','확인되지 않은 시장규모보다 검증할 고객·기술·중단 기준을 먼저 합의',[],true);
blocks(s,[['집중 범위','독일·영국에서 설치 자동화와 열 유연성을 검증하고 네덜란드는 확장 시장으로 준비'],['투입 조건','7명 역할 배치, 제품 API·실측 데이터 접근, 현지 실증 파트너 확보'],['12주 뒤 판단','유료 구매 의향·설치 공수·추가 순편익으로 확대 또는 축소를 결정']],{dark:true,gap:133});
}
// 26
{
let s=slide('부록: 재사용·지역화·신규 연구 구분','Digital Twin과 예측 정확도 개선은 이번 신규 과제의 주제로 선정하지 않음',['meeting']);
table(s,[['구분','내용','확인·개발 과제'],['기존 기술 재사용 후보','PV·부하 예측, ESS 스케줄링\nHVAC·가전 제어, 기존 실증','완료 연도·모델·실측 근거 확인'],['국가별 지역화','요금 API, 계통 신호, 기기 목록\n인증·계약·정산 규칙','공식 데이터와 버전 관리'],['신규 연구','혼합 기기 자동 검증\n열 유연성 추정과 안전 제어','기존 제품 대비 시간·순편익 입증'],['확장 후보','백업 지속시간·전력 허브\nVPP 운영·정산','파트너·인증·계약 조건 확보']], [290,480,382],{height:377,size:23});
}
// 27
{
let s=slide('부록: 시장규모를 실제 매출로 좁히는 방법','현재 자료로 국가별 TAM과 확정 매출을 산출할 수 없으므로 입력 자료를 명시');
blocks(s,[['설치 가능한 시장','국가별 HP·PV 보유 주택 × 기술 호환 비율 × 설치 권한·요금 조건을 충족하는 비율'],['접근 가능한 시장','계약 설치업체 수 × 연간 설치 가능 건수 × 모듈 채택률. 기존 제품의 출하 계획과 중복 제거'],['실제 매출','검증된 설치 물량 × 순판매 단가 + 평균 유료 가동 호수 × 운영 단가. 반품·지원비 별도 반영']],{gap:132});
note(s,'사업부 요청 자료: 국가별 출하·설치 채널·단가·목표 마진·인증비. 총 ESS/HP 시장을 HEMS 매출로 사용하지 않음.');
}
// 28
{
let s=slide('부록: 실증에서 계산할 고객 순편익','요금 차이·계절·사용 패턴을 통제하고 기존 운전과 비교');
cols(s,[['고객 가치 계산','순편익 = 전기요금 절감 + 실제 정산금\n− 서비스료 − 추가 손실·열화 비용\n\n기준선: 동일 기기·동일 요금의 기존 운전\n날씨·재실 보정 및 교차 기간 비교\n\n절감률과 함께 온도 이탈·수동개입·\n배터리 사용량을 공개'],['설치업체 가치 계산','편익 = 절약한 공수 × 시간당 비용\n+ 피한 재방문 비용 − 모듈 추가 원가\n\n설치 / 설정 / 검증 시간을 나누어 측정\n\n지원 기기·주택 조건별로 결과 공개\n\n계약 전에는 “절감 보장” 대신\n달성 조건과 측정 방식을 합의']]);
}
// 29–30 sources
for(const [title,keys] of [['부록: 출처와 사실 확인 범위',['meeting','de','de14','uk','fr','nl','au','ca']],['부록: 출처와 경쟁 제품 확인',['uae','kr','za','vpp','sma','tesla','oct','sch']]]){
let s=slide(title,'상세 URL은 각 관련 슬라이드의 발표자 노트에도 수록',keys);
keys.forEach((k,i)=>{txt(s,`${String(i+1).padStart(2,'0')}  ${S[k][0]}`,69,216+i*48,1134,39,23,i%2?C.ink:C.teal);});
}

await fs.writeFile(B+'/sources.json',JSON.stringify(S,null,2));
await fs.writeFile(B+'/deck-outline.json',JSON.stringify(pages.map((p,i)=>({slide:i+1,title:p.title,refs:p.refs})),null,2));
await (await PresentationFile.exportPptx(P)).save(B+'/candidate.pptx');
for(let i=0;i<pages.length;i++){
 const blob=await P.export({slide:pages[i].s,format:'png',scale:1});
 await fs.writeFile(B+`/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
 const l=await pages[i].s.export({format:'layout'});await fs.writeFile(B+`/layout-${i+1}.json`,await l.text());
 console.log('Rendered',i+1);
}
const finalPath=W+'/output/Energy_Pack_국가별_PainPoint_미래신기술_20260930_최종.pptx';
const result=await finalizePresentation({workspaceDir:W,candidatePath:B+'/candidate.pptx',finalPath,pythonExecutable:PY,integrityValidatorPath:SK+'/container_tools/inspect_presentation_package_integrity.py',layoutValidatorPath:SK+'/container_tools/inspect_presentation_layout_geometry.py',layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...tableOwners.flatMap(n=>['--require-native-table-slide',String(n)])],requiredNativeTableOwnerSlides:tableOwners,fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:B+'/validation.json'});
console.log(JSON.stringify(result));



