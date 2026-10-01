// Energy Pack Gateway Technology deck (pptxgenjs). Source: output/Energy_Pack_Gateway_Technology_2027-2030.html
const pptxgen = require('pptxgenjs');
const path = require('path');
const OUT = process.argv[2] || path.resolve(__dirname, '../output/Energy_Pack_Gateway_Technology_2027-2030.pptx');

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.title = 'Energy Pack Gateway Technology 2027–2030';
const F = 'Malgun Gothic';
const C = { dark: '142233', ink: '18212B', blue: '1D4F91', blueSoft: 'E3EBF6', copper: 'B8641F', copperSoft: 'F6EADF',
  muted: '5B6773', line: 'D3DAE1', panel: 'F2F4F7', white: 'FFFFFF', ok: '2D6A4F', okSoft: 'E1EFE7', warn: '8A5A00', warnSoft: 'F7ECD2',
  bad: '9B2C2C', badSoft: 'F5E1E1', darkMuted: 'A9B8C8' };

let n = 0;
function T(s, text, x, y, w, h, o = {}) {
  s.addText(text, Object.assign({ x, y, w, h, fontFace: F, fontSize: 14, color: C.ink, margin: 0, valign: 'top', isTextBox: true }, o));
}
function slide(section, title, sub, notes, dark = false) {
  n++;
  const s = pres.addSlide();
  s.background = { color: dark ? C.dark : C.white };
  if (section) T(s, section, 0.6, 0.38, 9, 0.28, { fontSize: 11, color: C.copper, bold: true, charSpacing: 1 });
  T(s, title, 0.6, 0.66, 12.1, 0.62, { fontSize: 26, bold: true, color: dark ? C.white : C.ink });
  if (sub) T(s, sub, 0.6, 1.3, 12.1, 0.4, { fontSize: 13, color: dark ? C.darkMuted : C.muted });
  T(s, 'Energy Pack Gateway Technology  ·  2026.09.30', 0.6, 7.08, 8, 0.25, { fontSize: 9, color: dark ? C.darkMuted : C.muted });
  T(s, String(n).padStart(2, '0'), 12.1, 7.08, 0.63, 0.25, { fontSize: 9, color: dark ? C.darkMuted : C.muted, align: 'right' });
  if (notes) s.addNotes(notes);
  return s;
}
function hdr(t) { return { text: t, options: { bold: true, color: C.white, fill: { color: C.dark }, fontSize: 10 } }; }
function table(s, rows, colW, o = {}) {
  const body = rows.map((r, i) => i === 0 && o.header !== false ? r.map(c => typeof c === 'string' ? hdr(c) : c) :
    r.map((c, j) => typeof c === 'string' ? { text: c, options: { bold: j === 0 && o.boldFirst !== false, fill: { color: i % 2 ? C.white : C.panel } } } : c));
  s.addTable(body, Object.assign({ x: 0.6, y: 1.85, w: colW.reduce((a, b) => a + b, 0), colW, fontFace: F, fontSize: 9.5, color: C.ink,
    border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'top', margin: [3, 5, 3, 5], autoPage: false }, o.opts || {}));
}
function pill(s, text, x, y, kind) {
  const m = { hi: [C.okSoft, C.ok], mid: [C.warnSoft, C.warn], lo: [C.badSoft, C.bad], gt: [C.copperSoft, C.copper], blue: [C.blueSoft, C.blue] }[kind];
  const w = Math.max(0.7, 0.13 * text.length + 0.3);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.28, fill: { color: m[0] }, line: { color: m[0] }, rectRadius: 0.14 });
  T(s, text, x, y, w, 0.28, { fontSize: 9.5, bold: true, color: m[1], align: 'center', valign: 'middle' });
  return w;
}
function card(s, x, y, w, h, fill = C.panel) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill } });
}

// ---------- 1 Title ----------
{
  const s = slide('', '', '', '조사 기준일 2026-09-30. 원문 보고서: output/Energy_Pack_Gateway_Technology_2027-2030.html (출처 68건).', true);
  T(s, 'ENERGY PACK / HEMS 신사업·신기술 발굴', 0.8, 1.5, 11, 0.35, { fontSize: 13, color: C.copper, bold: true, charSpacing: 2 });
  T(s, 'Energy Pack 2027–2030\nGateway Technology', 0.8, 2.0, 11.5, 1.9, { fontSize: 44, bold: true, color: C.white });
  T(s, '유럽·북미·중동 국가별 Pain Point에서 출발한 시장 장벽 분석과 필수 기술 도출', 0.8, 4.05, 11.5, 0.5, { fontSize: 18, color: C.darkMuted });
  const tags = ['독일·영국·프랑스·네덜란드·스페인/이탈리아', '미국 · California · Texas', 'UAE · 사우디 · 카타르'];
  let x = 0.8; tags.forEach(t => { const w = 0.16 * t.length + 0.5; s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 5.0, w, h: 0.4, fill: { color: '22344A' }, line: { color: '22344A' }, rectRadius: 0.2 }); T(s, t, x, 5.0, w, 0.4, { fontSize: 11.5, color: C.white, align: 'center', valign: 'middle' }); x += w + 0.2; });
  T(s, '2026.09.30 기준 · 2025–2026 공식·규제기관·업계 자료 68건', 0.8, 5.75, 11, 0.3, { fontSize: 11, color: C.darkMuted });
}

// ---------- 2 Answer ----------
{
  const s = slide('결론 먼저', '가장 큰 장벽은 기술 성능이 아니라 “Integration Tax”', '향후 3–5년 글로벌 Energy Pack 성장을 막는 가장 큰 장벽',
    'Sources: SEIA OBBB(25D 종료), CPUC NBT, Rijksoverheid salderingsregeling, photovoltaique.info, SolarPower Europe 2025, Verivox(blackout-news), Solar Builder soft costs, heatload HEMS integration, MCS Foundation, DOE VPP Liftoff 2025.');
  card(s, 0.6, 1.85, 12.13, 1.35, C.copperSoft);
  T(s, '보조금과 Net Metering이 사라지면서 Energy Pack의 가치는 여러 기기를 계통 규칙과 요금에 맞춰 함께 움직일 때만 생긴다. 그런데 그 결합(설치·연동·검증·권한·정산) 비용과 불확실성이 결합으로 얻는 가치보다 크다.', 0.85, 2.0, 11.6, 1.1, { fontSize: 16.5, bold: true, color: C.ink, valign: 'middle' });
  const cols = [
    ['① 단일 기기 경제성 소멸', ['미국 25D 세액공제 2025년 말 종료', '캘리포니아 수출 보상 약 75%↓', 'NL 상계제도 2027.1.1 일괄 종료', 'FR 잉여 PV 매입가 4c€/kWh', 'EU 가정용 ESS 2년 연속 감소(−6%)']],
    ['② 남은 가치는 유연성', ['독일 동적요금: 자동화 없으면 고정요금보다 비쌈 (6개월 604€ 대 481€)', '§14a 할인·VPP·자가소비 모두 PV·ESS·EV·HP 통합 제어가 전제']],
    ['③ 통합 비용이 해결 안 됨', ['미국 PV+ESS 비용의 58%가 soft cost', '독일 HP+PV HEMS 통합 2–8k€', '영국 MCS HP 설치자 약 9,000명', '북미 DER 중 VPP 등록 < 1/5']]];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 4.1;
    T(s, c[0], x, 3.45, 3.9, 0.35, { fontSize: 14, bold: true, color: C.blue });
    T(s, c[1].map((t, k) => ({ text: t, options: { bullet: { indent: 15 }, breakLine: k < c[1].length - 1 } })), x, 3.9, 3.9, 2.6, { fontSize: 12, paraSpaceAfter: 5 });
  });
  T(s, '→ 필요한 것은 “더 똑똑한 최적화”가 아니라 통합 비용을 0에 가깝게 만드는 기술', 0.6, 6.55, 12, 0.35, { fontSize: 14, bold: true, color: C.copper });
}

// ---------- 3 Stats ----------
{
  const s = slide('근거 수치', '통합 비용이 시장 확산을 막고 있다는 증거', '',
    'Solar Builder(soft cost 58%, 업계), heatload(2–8k€, 업계), MCS Foundation(9,000명), DOE VPP Liftoff 2025(<1/5), ScienceDirect smart thermostat review(12.9%), Ofgem 2025 survey(17%).');
  const st = [['58%', '미국 주택 PV+ESS 비용 중 soft cost', 'Solar Builder †'], ['2–8k€', '독일 HP+PV HEMS 통합 비용', 'heatload †'], ['9,000명', '영국 MCS 인증 HP 설치자 (2030년 연 45만 대 목표)', 'MCS Foundation'],
    ['< 1/5', '북미 DER 중 VPP에 등록된 용량', 'US DOE 2025'], ['12.9%', '여름 DR 이벤트 중 고객 override 비율', 'ScienceDirect 2025'], ['17%', '영국 TOU 미가입자 중 전환 의향', 'Ofgem 2025 조사']];
  st.forEach((a, i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.9 + Math.floor(i / 3) * 2.45;
    card(s, x, y, 3.9, 2.2);
    T(s, a[0], x + 0.3, y + 0.25, 3.4, 0.9, { fontSize: 40, bold: true, color: i < 3 ? C.copper : C.blue });
    T(s, a[1], x + 0.3, y + 1.2, 3.4, 0.6, { fontSize: 12.5 });
    T(s, a[2], x + 0.3, y + 1.8, 3.4, 0.25, { fontSize: 9.5, color: C.muted });
  });
  T(s, '† 업계·2차 자료 — 방향성 근거로 사용', 0.6, 6.8, 8, 0.25, { fontSize: 9.5, color: C.muted });
}

// ---------- 4 Gateway overview ----------
const GT = [
  ['GT-1', 'Zero-Touch Commissioning', '혼합 제조사 기기 자동 발견 → 능동 시험 신호로 배선·상·CT 방향·제어권 확인 → 국가 규칙 설정·준공 시험 보고서 자동 생성', '설치 인력 부족 · 설치·연동 비용 · 설치 오류로 인한 성능 저하', 'T1 + T2'],
  ['GT-2', 'Grid-Envelope Power Hub', '계량기급 측정 + 회로 단위 스위칭 HW + §14a·G100·DOE·100A 패널 한도를 실시간 지키며 EV·HP·ESS에 전력 배분', '기존주택 전기용량 부족 · 계통 혼잡 · 연계 대기', 'T4 + T5'],
  ['GT-3', 'Comfort-Bounded Thermal Flex', 'HP·에어컨·급탕을 쾌적 보장 계약이 붙은 가상 배터리로. 집마다 조절 가능한 kW·시간을 신뢰구간과 함께 산출', '최대 유연 부하(냉난방) 미활용 · 고객 override · 배터리 없는 가정 ROI', 'T3'],
  ['GT-4', 'Multi-Party Control & Verified Value', '고객·소매사·VPP·DSO·OEM 제어권 충돌을 계약 우선순위로 중재 + 반사실 기준선으로 실제 순편익 검증·정산', 'ROI 불확실 · 수익 스태킹 불가 · 공급사 lock-in · 신뢰 부족', 'T7 + T8']];
{
  const s = slide('결론', '반드시 확보해야 할 Gateway Technology 4개', 'Resilience Autopilot(T6)은 TX·CA·ES 지역 모듈로 분류 · Digital Twin/예측 정확도/대시보드/챗봇은 제외');
  GT.forEach((g, i) => {
    const x = 0.6 + (i % 2) * 6.13, y = 1.9 + Math.floor(i / 2) * 2.55;
    card(s, x, y, 5.93, 2.35);
    pill(s, g[0], x + 0.3, y + 0.25, 'gt');
    T(s, g[4], x + 4.5, y + 0.25, 1.15, 0.28, { fontSize: 10, color: C.muted, align: 'right', valign: 'middle' });
    T(s, g[1], x + 0.3, y + 0.62, 5.3, 0.4, { fontSize: 17, bold: true, color: C.blue });
    T(s, g[2], x + 0.3, y + 1.05, 5.35, 0.75, { fontSize: 11.5 });
    T(s, '해소 장벽: ' + g[3], x + 0.3, y + 1.85, 5.35, 0.4, { fontSize: 10, color: C.muted });
  });
}

// ---------- 5 Previous review ----------
{
  const s = slide('이전 산출물 검토', 'output 폴더 기존 PPT에서 이어받은 것과 바꾼 것', '');
  const L = ['“설치 자동화”·“열(냉난방) 유연성”을 핵심 과제로 본 판단 → 이번 조사에서도 가장 강한 근거 (GT-1, GT-3)', '경쟁사에 기능이 없다고 단정하지 않는 태도, 목표 수치를 “제안값”으로 구분', 'LLM이 요금·규제를 해석해 기기에 직접 명령하지 않는 안전 원칙 (T2에 반영)'];
  const R = ['범위 밖 한국·호주·남아공 제외 → Texas·사우디·카타르·스페인/이탈리아 추가', '신기술 후보 6개 → 12개, 각 10개 질문에 답변', '경쟁사 4곳 → 24곳, “아직 못 푼 부분” 중심', '근거 없는 매출 시나리오(작성자 가정) 삭제 → Gateway Technology·가치사슬로 결론 구성', '새 장벽 추가: 계통 연계 용량(패널·§14a·G100), 다자간 제어권'];
  [[L, '유지', C.ok, C.okSoft], [R, '보완·변경', C.copper, C.copperSoft]].forEach((c, i) => {
    const x = 0.6 + i * 6.13;
    card(s, x, 1.9, 5.93, 4.9, c[3]);
    T(s, c[1], x + 0.35, 2.1, 5, 0.4, { fontSize: 18, bold: true, color: c[2] });
    T(s, c[0].map((t, k) => ({ text: t, options: { bullet: { indent: 15 }, breakLine: k < c[0].length - 1 } })), x + 0.35, 2.65, 5.3, 4, { fontSize: 13, paraSpaceAfter: 9 });
  });
}

// ---------- 6-9 Country tables ----------
const CH = ['국가', '시장/제도 변화', '고객/Utility Pain Point', '현재 해결방식', '아직 남은 Gap', '필요한 신기술'];
const CW = [1.05, 3.25, 3.35, 1.55, 1.85, 1.08];
const countries = [
  ['유럽 ①', '독일·영국: 규제는 앞서가는데 현장 설치·연동이 따라가지 못함', [
    ['독일\nDE', '• 2025년부터 모든 공급사 동적요금 의무\n• §14a: 4.2kW 초과 HP·Wallbox·ESS는 DSO 제어 대상. Modul 3(시간가변 망요금) 2025.4 의무. 연 110–190€~최대 360€ 절감\n• Solarspitzengesetz(2025.2): 음가격 시 보상 없음, 스마트미터 없으면 송출 60% 제한\n• 가정요금 38.69€/100kWh (EU 2위)\n• HP 29.9만 대(+50%, 80% 기존주택) · 가정 ESS 36.6만 대(−20%)\n• 2026년부터 V2G 이중 망요금 폐지',
      '• 고객: 동적요금 가입해도 자동 제어 없으면 손해 (604€ 대 481€)\n• 고객: 음가격·60% 제한으로 PV 수익 감소, ESS 운전 판단 불가\n• 설치업체: EEBus·SG-Ready·Modbus 혼용 → 게이트웨이 누락 오류, 복합 기술 인력 부족\n• DSO: 스마트미터 5.5%(5,650만 중 310만), §14a 기준정보 프로세스 미비',
      'SMA, 1KOMMA5 Heartbeat(6만 가구·1GW), Enpal, sonnen VPP, gridX, HP SG-Ready', '• 타사 기기 자동 연동·검증\n• §14a·동적요금·60%·음가격 동시 준법 증명\n• HP 열 유연성 활용', 'T1 · T2\nT4 · T3\nT10'],
    ['영국\nUK', '• 2026.10 가격상한 26p/kWh, 기본료 54.8p/일\n• MHHS: 2027.5까지 30분 정산 이행\n• Warm Homes Plan: 2030년 HP 연 45만 대 목표, 2026 상반기 설치 −22%\n• 2026.8.27 플러그인 태양광(800VA) 합법화\n• DFS 참여 160만 가구',
      '• 공급 병목: MCS HP 설치자 약 9,000명, 목표 달성에 설치 속도 5배 필요\n• 고객: TOU 절감 기대 49% vs 미가입자 전환 의향 17%\n• 고객: 높은 기본료로 절감 체감 약함, 일부 스마트미터 비작동\n• DNO: G100 송출 한도(5초 대응) 요구로 연계 지연',
      'Octopus Kraken(2GW)·Intelligent Octopus, Tesla, GivEnergy', '• 공급사 tariff에 묶인 자동화\n• 설치 1건당 시간 단축 기술 부재\n• 임차인·공동주택용 부재', 'T1 · T3\nT8 · T12\nT11']],
    'BNetzA 동적요금·§14a, Vattenfall §14a Modul, pv magazine Solarspitzen·배터리, Eurostat 2025, BWP HP, Mobility House V2G, Utility Partners, Verivox, heatload; Ofgem price cap·2025 survey, Elexon MHHS, GOV.UK HP stats, MCS Foundation, Fuse Energy plug-in solar, NESO DFS, Reonic G100.'],
  ['유럽 ②', '프랑스·네덜란드·남유럽: 요금·보상 구조가 바뀌어 기존 기기 운전이 어긋남', [
    ['프랑스\nFR', '• Linky 95% (3,800만)\n• 2025.11–2027: 약 1,100만 고객 Heures Creuses를 낮 시간대로 순차 변경\n• 2025.3 잉여 PV 매입가 12.69→4c€/kWh, 자가소비 프리미엄 반감\n• MaPrimeRénov 2026.1–2 일시 중단',
      '• 고객: 전기온수기·축열난방이 옛 심야 타이머에 고정, 새 시간대와 어긋남\n• 고객: 수출 가치가 낮아 자가소비율이 경제성 좌우, 제어할 기기는 비연결\n• 시장: 보조금 중단·재개로 HP·리노베이션 수요 흔들림',
      'Enedis 신호, EDF Tempo, Schneider Wiser', '• 비연결 기존 온수기·난방기 제어\n• 요금 개편 자동 반영·검증', 'T2 · T9\nT3'],
    ['네덜란드\nNL', '• 2027.1.1 상계제도 일괄 종료, 2027–2030 송출 보상 공급가의 50% 이상\n• 공급사별 송출비용(terugleverkosten)\n• 광범위한 망 혼잡',
      '• 고객: PV 수익 급락, ESS 회수기간 > 보증기간\n• DSO: 신규 HP·EV 연계 대기, 한낮 송출 과잉', 'ESS 판매 증가, 동적요금', '• 배터리 없이 잉여 PV 흡수(급탕·예열)\n• 혼잡 보상 연계 한도 제어', 'T3 · T4\nT8'],
    ['스페인·이탈리아\nES·IT', '• 2025.4.28 이베리아 대정전(ENTSO-E 최종보고) 이후 가정용 배터리 수요 급증\n• 이탈리아 도매가 2025년 €115.9/MWh (FR €61.1)',
      '• ES: PV 있어도 grid-forming·islanding 없어 정전 시 사용 불가\n• IT: 높은 요금, 3kW 저계약용량으로 HP·EV 동시 사용 불가', 'Huawei·SolarEdge·Tesla 백업 인버터', '• 정전 시 중요부하 생존 제어\n• 저계약용량 내 전력 배분', 'T6 · T4\nT5']],
    'Enedis 2025.10.15, photovoltaique.info, Selectra Linky, CAPEB MaPrimeRénov; Rijksoverheid, gridX; ENTSO-E Iberian final report, The Local Italy 2026.9.'],
  ['북미', '미국·California·Texas: 보조금 종료 후 패널 용량·수익 불확실·정전이 장벽', [
    ['미국 전체\nUS', '• 25D(소유형 30%) 2025.12.31 종료, 리스·PPA는 48E 유지 → TPO로 이동\n• 가정요금 18.02¢(2026) 상승, 데이터센터 부하 증가\n• VPP 33GW, DOE 목표 2030년 80–160GW',
      '• 고객: 단독주택 29.5%가 100A 이하 패널, 서비스 증설 $7k–13k(업계)\n• 설치업체: 비용 58%가 soft cost(영업·설계·허가·연계)\n• Utility: DER 늘었지만 VPP 등록 <1/5, OEM 사일로',
      'SPAN·Schneider Pulse·ConnectDER, Tesla/Enphase/SolarEdge, Sunrun TPO', '• 제조사 중립 패널 증설 회피(PCS)\n• 한 가정의 다중 VPP·프로그램 동시 운영', 'T5 · T1\nT7 · T8'],
    ['California\nUS-CA', '• NBT: 수출 보상 약 75%↓, ESS attach 약 70%\n• PG&E 평균 39.25¢/kWh, TOU 16–21시\n• Rule 21: IEEE 2030.5/CSIP 의무\n• DSGS VPP 540MW 시험 성공, 예산 삭감 후 일부 복원',
      '• 고객: 자가소비·백업·VPP가 같은 배터리 잔량을 두고 충돌\n• 고객: VPP 프로그램이 예산 따라 사라져 수익 예측 불가\n• Utility: PSPS·산불 정전', 'Tesla Storm Watch, Sunrun CalReady, Enphase', '• 프로그램 변경 시 자동 재구성되는 수익 스태킹\n• 백업·VPP 위험 기반 배분', 'T6 · T7\nT8 · T3'],
    ['Texas\nUS-TX', '• ERCOT ADER 파일럿 → 정식 규칙화, VPP 3개 25.5MW 승인\n• 소매 자유화: Tesla Electric Powerwall 2대 월 $35 리스\n• 대형부하(데이터센터) 급증',
      '• 고객: Uri(2021) 450만 가구 최장 4일, Beryl(2024) 260만 고객 수일 정전(폭염)\n• 고객: 폭염 정전 시 ESS만으로 에어컨 장시간 불가\n• TBPP는 중요시설 대상, 가정 지원 아님', 'Tesla Electric VPP, Sunrun/NRG, sonnen/SOLRITE', '• 냉방 포함 생존시간 극대화\n• 소매사·ERCOT·TDU 신호 중재', 'T6 · T3\nT7 · T5']],
    'SEIA OBBB, EIA/Utility Dive, DOE VPP Liftoff 2025, ACEEE panel survey, Solar Builder; CPUC NBT, Solar.com PG&E, CPUC Rule 21, Brattle VPP test, pv magazine USA DSGS; Utility Dive Tesla ERCOT, TPPF Uri, Texas Tribune Beryl, Electrek, Texas Energy Fund TBPP.'],
  ['중동', 'UAE·사우디·카타르: 소비자 ROI가 아니라 Utility·국가의 냉방 피크 문제', [
    ['UAE\nAE', '• DEWA 누진 23–38 fils/kWh + 연료할증 6 fils(2026.7)\n• Shams Dubai net metering 725MW, 8,430개 건물(2025.6)\n• Dubai DSM 2050: 2030년 30% 절감, Efficient Cooling',
      '• 빌라 고객: 여름 냉방비가 누진 상위 구간, 다수 split AC 비연결\n• 가정용 TOU 부재 → ESS 소비자 ROI 약함\n• Utility: 냉방 피크 저감 목표, 아파트는 지역냉방이라 제어 대상 아님', 'DEWA 앱·스마트미터, AC 제조사 앱', '• 기존 AC 원격 제어·능력 추정\n• Utility·개발사 부담 사업모델', 'T9 · T3\nT8'],
    ['사우디\nSA', '• 가정요금 0.18 SAR/kWh(월 6,000kWh 이하), 초과 0.30\n• 2020 소규모 PV 규정(net billing) 있으나 보급 미미\n• 가정 부문이 전력의 48%',
      '• SEC·정부: 여름 피크 65GW+ (겨울 약 2배), 에어컨이 가정 소비 60%+·피크 약 70%(업계 추정)\n• 고객: 낮은 요금으로 PV·ESS 비경제적, 절감 동기 약함', '효율 기준 강화, 신도시 스마트홈', '• Utility·국가 가치로 과금하는 모델\n• 대단지 냉방 flex 집단 제어', 'T3 · T9\nT7'],
    ['카타르\nQA', '• Kahramaa BeSolar 분산 PV net billing, Tarsheed 효율 프로그램', '• 사우디와 유사한 저요금·냉방 피크 구조, 시장 규모 작음', 'Kahramaa 프로그램', '사우디와 동일', 'T3 · T9']],
    'DEWA slab tariff, SurgePV Shams Dubai†, Dubai SCE DSM 2050; SERA consumption tariff, ScienceDirect 2025 Saudi PV incentives, vision2030.ai†; Kahramaa BeSolar.']];
countries.forEach(c => {
  const s = slide('국가별 Pain Point · ' + c[0], c[1], '', 'Sources: ' + c[3]);
  table(s, [CH, ...c[2]], CW, { opts: { fontSize: 11.5 } });
});

// ---------- 10 Barriers chart ----------
{
  const s = slide('Gap 분석', '기술은 이미 있는데 왜 시장이 안 크는가', '12개 국가·State Pain Point를 원인별로 분류 — 기술 “성능” 부족은 거의 없고, 기술이 현장에 닿는 경로에서 막힘',
    '저자가 2장 표의 Pain Point를 분류해 센 값(12 = DE, UK, FR, NL, ES/IT, US, CA, TX, AE, SA, QA + EU 공통). 정량 조사 결과가 아님.');
  const labels = ['비연결 기존 기기', '정전 Resilience', 'Comfort·신뢰(override)', '제어권·이해관계자 분절', '계통 연계 용량 제약', '설치·연동 복잡성', 'ROI 불확실·보조금 소멸'];
  const vals = [4, 4, 6, 7, 7, 9, 10];
  s.addChart(pres.charts.BAR, [{ name: '국가 수', labels, values: vals }], {
    x: 0.6, y: 1.9, w: 7.2, h: 4.9, barDir: 'bar', chartColors: [C.blue], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 11, dataLabelColor: C.ink,
    catAxisLabelFontFace: F, catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, valAxisHidden: true, valAxisMaxVal: 12, valAxisMinVal: 0,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false, showTitle: true, title: '근본 장벽이 나타난 국가 수 (12개 중)', titleFontFace: F, titleFontSize: 12, titleColor: C.muted });
  card(s, 8.2, 1.9, 4.53, 4.9);
  T(s, '해석', 8.5, 2.1, 4, 0.35, { fontSize: 15, bold: true, color: C.copper });
  T(s, [
    { text: 'PV·Load Forecast, ESS·TOU 최적화, EV·HVAC 스케줄링은 이미 상용 기능', options: { bullet: { indent: 15 }, breakLine: true } },
    { text: '상위 장벽은 ROI 불확실과 설치·연동 복잡성 → 모두 “통합 비용” 문제', options: { bullet: { indent: 15 }, breakLine: true } },
    { text: '계통 연계 용량과 제어권 분절은 이전 검토에서 빠져 있던 새 장벽', options: { bullet: { indent: 15 }, breakLine: true } },
    { text: '정전·비연결 기기는 특정 지역(TX·CA·ES / GCC·FR) 장벽', options: { bullet: { indent: 15 } } }], 8.5, 2.6, 4, 4, { fontSize: 12.5, paraSpaceAfter: 10 });
}

// ---------- 11 Why unsolved ----------
{
  const s = slide('Gap 분석', '장벽별로 기존 HEMS가 풀지 못하는 이유', '', 'Sources: Solar Builder, heatload, MCS Foundation, Verivox, SolarPower Europe, gridX, ACEEE, pv magazine Solarspitzen, Reonic G100, DOE VPP Liftoff, pv magazine USA DSGS, ScienceDirect thermostat review, Ofgem 2025, Enedis, Texas Tribune, ENTSO-E.');
  table(s, [['근본 장벽', '왜 기존 HEMS로 안 풀리나', '대표 증거'],
    ['설치·연동 복잡성', 'HEMS는 자사 생태계 기기만 깊게 제어. 타사 기기는 사람이 프로토콜·레지스터·배선을 맞춰야 하고, 설치자는 부족하며 설정 오류는 성능 저하로 이어짐', 'US soft cost 58%, DE 통합 2–8k€, UK 설치자 9,000명, 미커미셔닝 HP SCOP 15–30%↓(업계)'],
    ['ROI 불확실', '최적화는 “예상 절감”만 보여줌. 반사실 기준선으로 검증된 절감이 없어 금융·보증·성과형 가격 불가. 보조금이 빠지자 판매 감소', 'DE 동적요금 역전, EU 가정 ESS −6%, NL 회수기간 > 보증'],
    ['계통 연계 용량', 'HEMS는 요금 최적화 SW라 연계점 전력 한도를 법적 효력 있게 “보장”하지 못함 → 패널 증설·연계 대기가 먼저 발생', 'US 29.5% ≤100A, §14a·60% 제한, G100, NL 혼잡'],
    ['제어권 분절', '한 가정 배터리를 소매사·VPP·DSO·OEM 클라우드가 각자 제어하려 함. 중재 계층이 없어 한 곳만 등록하거나 충돌', 'VPP 등록 <1/5·OEM 사일로, DSGS 예산 변동'],
    ['Comfort·신뢰', '냉난방은 최대 유연 부하인데, 쾌적 한계를 모르고 제어하면 고객이 바로 override', 'DR override 12.9%(대부분 2–3시간 내), UK TOU 전환 의향 17%'],
    ['비연결 기존 기기', 'HEMS는 API 있는 기기만 다룸. 중동 split AC, 프랑스 온수기, 미국 구형 HVAC는 제어 대상에서 빠짐', 'UAE 빌라 냉방, FR ballon 타이머'],
    ['정전 Resilience', '백업 모드가 “배터리 100% 충전”에 그침. 부하 단위 생존시간 최적화·냉방 조정 불가', 'TX Uri·Beryl, ES 대정전']], [2.2, 6.2, 3.73], { opts: { fontSize: 12.5 } });
}

// ---------- 12 Hypotheses ----------
{
  const s = slide('가설 A–E 검증', '이미 상용화된 부분과 남은 부분', '', 'Sources: 1KOMMA5 Heartbeat, Tesla Storm Watch, Samsung SmartThings, Canary Media SPAN, EcoFlow OCEAN Pro.');
  const P = (t, k) => ({ text: t, options: { bold: true, color: { hi: C.ok, mid: C.warn, lo: C.bad }[k], fill: { color: { hi: C.okSoft, mid: C.warnSoft, lo: C.badSoft }[k] } } });
  table(s, [['가설', '이미 상용화된 것', '비판', '판정', '재정의'],
    ['A. Region-Adaptive Energy AI', '동적요금 API·국가별 요금 연동은 흔함 (Kraken, Heartbeat AI, gridX, SMA)', '“AI가 규제를 이해”만으로는 상품이 아님. 오해석 시 준법 리스크. 차별화는 검증된 규칙 변환과 준법 증명', P('조건부', 'mid'), 'T2 Regulation-as-Code (GT-1 구성요소)'],
    ['B. Auto-Commissioning', '자사 생태계 내 앱 커미셔닝 (Tesla, Enphase, Huawei, Sigenergy)', '혼합 제조사·기존 설치 기기·능동 시험으로 배선·상·CT 방향 검증은 공백. 시장 장벽 1위와 직결', P('채택', 'hi'), 'T1 → GT-1'],
    ['C. Energy Autopilot', 'Tesla Savings/Storm Watch, Intelligent Octopus, Heartbeat AI, Samsung AI Energy Mode가 이미 제공', '단독으로는 신기술 아님. 남은 것은 냉난방 쾌적 보장과 다자간 권한 충돌 처리', P('단독 불가', 'lo'), 'T3·T7에 흡수'],
    ['D. Resilience / VPP', 'Storm Watch, CA·TX VPP, sonnen VPP', 'VPP와 백업이 같은 잔량을 두고 경쟁. 부하 단위 생존 최적화·폭염/한파 냉난방 조정은 공백. 수요는 TX·CA·ES 집중', P('지역 특화', 'mid'), 'T6 + T7'],
    ['E. Plug&Play Power Hub', '미국: SPAN, Schneider Pulse, ConnectDER, Tesla Backup Switch, EcoFlow OCEAN Pro', '미국 단상 패널형은 경쟁 치열. 유럽 3상·§14a·SG-Ready·G100 통합 허브, HP·AC 제조사 관점 허브는 공백', P('채택(EU 우선)', 'hi'), 'T4 + T5 → GT-2']], [2.1, 3.1, 4.1, 1.15, 1.68], { opts: { fontSize: 11.5 } });
  T(s, '추가 제안: T7 제어권 중재 · T8 검증 가치 정산 · T9 기존 기기 활성화 · T12 임차인용 Micro Pack — A–E에는 없었지만 조사에서 장벽으로 반복 등장', 0.6, 6.55, 12.1, 0.4, { fontSize: 12, bold: true, color: C.copper });
}

// ---------- 13 Competitor matrix ----------
{
  const s = slide('경쟁사', '누가 어디까지 하고 있나', '● 상용 제품  ◐ 일부 국가·파트너·제한적  · 없음/미확인 — 공개 제품 정보 기준, 국가·모델별 차이 있음',
    'Sources: Enphase IQ EM(2026.3), pv Europe Huawei, Sigenergy, sonnen VPP, ESS News 1KOMMA5, EV Infrastructure News Kraken, POWER Voltus–Octopus, Schneider Wiser, Canary Media SPAN, Samsung SmartThings, Mitsubishi §14a.');
  const D = { f: '●', h: '◐', n: '·' };
  const rows = [
    ['Tesla', 'ffnnhff'.split(''), 'f', '냉난방·타사 기기 제어 없음, 폐쇄 생태계'],
    ['Enphase', 'fffhnff'.split(''), 'h', 'IQ Energy Mgmt는 온수기·충전기 일부(2026.3 AU/NZ)'],
    ['SolarEdge', 'fffhnfh'.split(''), 'h', 'HP 열 유연성, 혼합 기기'],
    ['Huawei', 'fffhnfh'.split(''), 'h', 'HP는 SG-Ready·릴레이 수준, 일부 국가 보안 규제'],
    ['Sigenergy', 'fffnnhh'.split(''), 'h', 'DC 양방향 충전 통합 강점, 열부하 없음'],
    ['sonnen (Shell)', 'hfhnnff'.split(''), 'h', 'ESS 중심 VPP 250MWh, HP flex 약함'],
    ['SMA', 'fffhnfh'.split(''), 'h', '설정이 설치자 전문성에 의존'],
    ['1KOMMA5 · Enpal', 'ffffhff'.split(''), 'f', '가장 근접. 1GW·레트로핏 박스, 자사 채널·독일 중심, 쾌적 보장 없음'],
    ['Octopus / Kraken', 'hhffnff'.split(''), 'f', '유연성 SW 최강(2GW), 자사 요금 가입 전제, HW 설치 병목'],
    ['Schneider', 'hhfnnfn'.split(''), 'h', 'Pulse(US) 패널 통합, Wiser AI는 자사 조합 한정'],
    ['Siemens', 'nnfnnhn'.split(''), 'n', 'DSO·그리드 SW·충전 중심, 가정 HEMS 존재감 낮음'],
    ['Bosch', 'nnnffhn'.split(''), 'n', 'Energy Manager는 PV 잉여 연동 수준'],
    ['SPAN · ConnectDER', 'nhhnnfh'.split(''), 'h', '미국 전용 패널 교체형, Eaton·Landis+Gyr 제휴'],
    ['EcoFlow · Anker', 'fffnnhn'.split(''), 'h', '가격·올인원, 계통 준법·VPP 약함'],
    ['BYD', 'nffnnnn'.split(''), 'n', '셀·모듈 공급자, HEMS 없음'],
    ['Google Nest', 'nnnnfhf'.split(''), 'h', '온도조절기 DR만, 전력기기 없음'],
    ['Amazon', 'nnnnhhn'.split(''), 'n', 'Alexa 에너지 대시보드, Matter 허브'],
    ['Samsung', 'nhhfffh'.split(''), 'f', '가전 AI Energy Mode·SmartThings Energy, 계통 준법·전력 HW 약함'],
    ['Midea · Haier', 'nhnffhn'.split(''), 'h', '원가 경쟁력, HEMS·VPP 생태계 약함'],
    ['Daikin', 'nnnffhn'.split(''), 'n', 'SG-Ready·Modbus 연동은 외부 HEMS 의존'],
    ['Mitsubishi Electric', 'nnnffhn'.split(''), 'n', '§14a 대응 SG-Ready, 유연성 수익화 없음'],
    ['Panasonic', 'hhhffhn'.split(''), 'h', '일본 HEMS(AiSEG) 경험 해외 확장 미흡']];
  const cell = k => ({ text: D[k], options: { align: 'center', color: k === 'f' ? C.blue : k === 'h' ? C.copper : 'B8C0C8', fontSize: 11 } });
  const body = [['업체', 'PV', 'ESS', 'EV', 'HP', 'HVAC', 'HEMS', 'VPP', 'AI 운전', '아직 못 푼 부분'].map(hdr)];
  rows.forEach((r, i) => body.push([{ text: r[0], options: { bold: true, fill: { color: i % 2 ? C.panel : C.white } } }, ...r[1].map(cell), cell(r[2]), { text: r[3], options: { fill: { color: i % 2 ? C.panel : C.white } } }]));
  s.addTable(body, { x: 0.6, y: 1.8, w: 12.13, colW: [1.8, 0.55, 0.55, 0.55, 0.55, 0.6, 0.6, 0.55, 0.7, 5.68], fontFace: F, fontSize: 9, color: C.ink, rowH: 0.215,
    border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [1, 5, 1, 5], autoPage: false });
}

// ---------- 14 Competitor gaps ----------
{
  const s = slide('경쟁사', '경쟁사가 아직 풀지 못한 세 가지 공백', '');
  const g = [['공백 1', '냉난방 기기 회사에는 전력·계통 스택이 없다', 'Daikin·Mitsubishi·Bosch·Midea는 HP를 팔지만 유연성 수익화는 외부 HEMS에 넘긴다.'],
    ['공백 2', '전력기기 회사는 열부하를 모른다', 'Tesla·Enphase·Huawei·Sigenergy는 ESS 중심. 가장 큰 유연 부하인 냉난방을 쾌적 보장과 함께 제어하지 못한다.'],
    ['공백 3', '누구도 “중립 통합자”가 아니다', '1KOMMA5·Octopus가 가장 앞서지만 자사 설치 채널이나 요금 가입이 전제. 제조사 중립 커미셔닝·권한 중재는 비어 있다.']];
  g.forEach((a, i) => {
    const x = 0.6 + i * 4.1;
    card(s, x, 1.9, 3.9, 3.2);
    pill(s, a[0], x + 0.3, 2.15, 'blue');
    T(s, a[1], x + 0.3, 2.6, 3.35, 0.8, { fontSize: 15.5, bold: true });
    T(s, a[2], x + 0.3, 3.5, 3.35, 1.5, { fontSize: 12, color: C.muted });
  });
  card(s, 0.6, 5.4, 12.13, 1.3, C.copperSoft);
  T(s, [{ text: '우리가 노릴 위치  ', options: { bold: true, color: C.copper } }, { text: 'HP·에어컨을 직접 만들면서 전력 HW(Energy Pack)와 계통 준법 SW를 함께 가진 회사는 공백 1과 2를 동시에 메울 수 있다.', options: { bold: true } }],
    0.9, 5.5, 11.6, 1.1, { fontSize: 15, valign: 'middle' });
}

// ---------- 15-16 Tech tables ----------
const TH = ['신기술', '해결 Pain Point', 'Target', '기존 기술과 차이', 'AI', 'HW', '핵심 기술', '경쟁사 현황', '차별화', '필요성'];
const TW = [1.7, 1.25, 1.0, 2.2, 1.1, 1.1, 1.25, 1.2, 0.66, 0.67];
const RP = (t, k) => ({ text: t, options: { bold: true, fontSize: 8.5, color: { hi: C.ok, mid: C.warn }[k], fill: { color: { hi: C.okSoft, mid: C.warnSoft }[k] } } });
const techRows = [
  ['T1 Active-Probe Auto-Commissioning', '설치 인력·시간·오류', 'DE UK US NL', '“목록에서 기기 선택”이 아니라 시험 신호로 실제 배선·응답 확인', '시스템 식별, 이상 판정', '멀티프로토콜 GW + CT 입력', '능동 perturbation 식별, 자동 준공 시험', '자사 생태계 내만', RP('높음', 'hi'), RP('매우 높음', 'hi')],
  ['T2 Regulation-as-Code Compiler', '국가별 요금·규제 변경', '전 국가', '요금 API 연동을 넘어 규칙을 검증 가능한 제약으로 컴파일·버전 관리', '문서→규칙 후보 추출', '불필요', '형식 검증, 정책팩 배포', '수작업 지역화', RP('중간', 'mid'), RP('높음', 'hi')],
  ['T3 Comfort-Bounded Thermal Flex', '냉난방 flex 미활용, override', 'DE UK NL FR TX CA SA AE', '스케줄링이 아니라 가정별 쾌적 한계·flex 용량을 신뢰구간으로 보증', '열 응답 식별, 쾌적 학습', '실내 센서(습도)', '확률적 envelope, 쾌적 계약', 'Octopus(HP), Nest(DR) 일부', RP('높음', 'hi'), RP('매우 높음', 'hi')],
  ['T4 Grid-Envelope Controller', '§14a·G100·DOE·계약용량', 'DE UK NL IT US', '요금 최적화가 아니라 연계점 한도를 법적 수준으로 보장하는 실시간 배분', '우선순위·손실 최소 배분', '계량기급 측정, 빠른 제어', '<1초 제어 루프, 준법 로그', '기기별 스로틀 수준', RP('높음', 'hi'), RP('매우 높음', 'hi')],
  ['T5 Virtual Service Upgrade Hub', '패널 증설 비용', 'US IT UK', '패널 교체가 아닌 retrofit 허브, 제조사 중립', '부하 예측 기반 차단 계획', '핵심 (PCS, 릴레이)', 'UL 3141 PCS, 미터칼라', 'SPAN·Schneider·ConnectDER', RP('중간', 'mid'), RP('높음', 'hi')],
  ['T6 Resilience Autopilot', '정전 생존', 'TX CA ES', '“100% 충전”이 아니라 부하 단위 생존시간 극대화·냉방 사이클링', '정전 위험·부하 우선순위', '회로 차단, grid-forming', '생존시간 최적화', 'Tesla Storm Watch', RP('중간', 'mid'), RP('지역 높음', 'mid')],
  ['T7 Control-Rights Arbiter', '다자간 제어 충돌', 'US UK DE CA TX', '단일 VPP 등록이 아니라 계약 우선순위 기반 다중 참여', '기회비용 기반 배분', '로컬 실행 GW', '권한 원장, 로컬 중재', '공백', RP('높음', 'hi'), RP('높음', 'hi')],
  ['T8 Verified Savings & Settlement', 'ROI 불확실', '전 국가', '예상 절감이 아니라 반사실 기준선 기반 검증 절감·성과형 과금', '반사실 기준선', '계량 데이터', 'M&V, 감사 가능 정산', 'Octopus 일부, 대부분 추정치', RP('중간', 'mid'), RP('높음', 'hi')],
  ['T9 Legacy Appliance Enablement', '비연결 AC·온수기', 'AE SA QA FR US', 'API 없는 기기를 NILM + retrofit 액추에이터로 제어', '부하 분리·능력 추정', 'IR·릴레이 모듈', '폐루프 확인 제어', '단순 스마트플러그', RP('중간', 'mid'), RP('중간', 'mid')],
  ['T10 Mobility-Aware V2H', 'EV를 가정 배터리로', 'DE US UK', 'OEM별 V2H가 아니라 ISO 15118-20 중립 + 이동 필요량 보증', '출발·주행 예측', 'DC 양방향 결합', 'DC 버스 공유', 'Sigenergy, GM, Ford', RP('중간', 'mid'), RP('2028+', 'mid')],
  ['T11 Remote Performance Assurance', '설치 후 성능 저하·방문 비용', 'UK DE', '알림 모니터링이 아니라 원격 자동 교정(히트커브 등)', '이상 원인 진단', '기존 GW', '자동 재튜닝', '제조사 원격 진단 일부', RP('중간', 'mid'), RP('중간', 'mid')],
  ['T12 Renter Micro Pack', '임차인·공동주택 배제', 'UK DE NL FR', '설치공사 없는 800VA 플러그인 + 소형 ESS + 가전 연동', '가전 협조 제어', '플러그인 ESS', '역송 제한 준수', 'EcoFlow, Anker, Zendure', RP('중간', 'mid'), RP('중간', 'mid')]];
[[0, 6, 'T1–T6'], [6, 12, 'T7–T12']].forEach(([a, b, lab]) => {
  const s = slide('신기술 후보 12개', '평가표 (' + lab + ')', '');
  table(s, [TH, ...techRows.slice(a, b)], TW, { opts: { fontSize: 12 } });
});

// ---------- 17-28 Tech details ----------
const Q = ['1. 출발 Pain Point', '2. 기존 HEMS로 어려운 이유', '3. 핵심 기술', '4. AI가 필요한 이유', '5. HW 필요성', '6. 고객 Value', '7. Utility/Grid Value', '8. 경쟁사', '9. 2–3년 격차 포인트', '10. 특허 아이디어'];
const details = [
  ['T1', 'Active-Probe Auto-Commissioning', 'GT-1', '혼합 제조사 기기를 연결하면 시험 신호로 배선·상·제어권을 확인하고 준공 보고서까지 자동 생성', [
    '독일 HEMS 통합 2–8k€·프로토콜 혼용 오류, 영국 설치자 9,000명, 미국 soft cost 58%',
    '기기 정보를 설치자 입력에 의존. CT 방향 반전·상 오결선·SG-Ready 릴레이 극성·제어 권한 미부여가 운전 후 성능 저하로 드러남',
    '① 멀티프로토콜 탐색(EEBus, SunSpec Modbus, OCPP, Matter 1.4, SG-Ready) ② 작은 시험 신호(ESS 1kW, HP ±1K, EV 6A↔10A) + 계량 응답으로 배선·상·지연·제어 방향 식별 ③ 국가 규칙 설정 자동 생성 ④ 준공 시험 보고서(§14a, G100, Rule 21) 자동 발급',
    '여러 기기 응답이 겹친 계량 신호를 기기별로 분리하는 시스템 식별 문제. 규칙 기반으로는 조합 폭발',
    '필요. 상별 CT 입력·고정밀 동기 계측, RS485·Ethernet·Thread/Wi-Fi, 릴레이 출력을 갖춘 Energy Pack 게이트웨이',
    '설치 당일부터 제대로 작동, 재방문 감소, 설치비 절감',
    '연계 기기의 준법 상태가 증명된 채 등록, §14a·G100 제어가 실제로 동작',
    'Tesla·Enphase·Huawei·Sigenergy는 자사 기기 앱 커미셔닝만. 1KOMMA5 Heartbeat Box는 타사 연결하나 능동 검증 미확인',
    '자사 HP·AC 내부 상태 접근 + 현장에서 쌓이는 “기기 모델별 응답 서명” 라이브러리',
    '(a) 기기별 시험 신호를 시간·주파수 직교 설계해 여러 기기의 배선·상을 동시 식별 (b) 식별 결과 + 국가 규칙팩으로 준공 시험 절차 자동 생성·판정']],
  ['T2', 'Regulation-as-Code Compiler', 'GT-1 구성요소', '국가·DSO·공급사별 요금·규제를 검증 가능한 제약으로 컴파일해 서명된 정책팩으로 배포', [
    'FR Heures Creuses 개편(1,100만 고객), DE §14a·Solarspitzen, NL 2027 상계 종료, CA DSGS 예산 변동 — 규칙이 자주·지역마다 바뀜',
    '규칙을 사람이 코드로 옮김 → 반영 지연, 오류는 고객 손해·준법 위반',
    '요금표·망규정·약관을 기계 판독 제약(시간대 단가, 한도, 금지 조건)으로 변환 → 모순·누락 형식 검증 → 서명된 정책팩 배포. 청구서·스마트미터 데이터로 가입 요금 역추정',
    '비정형 문서(PDF 요금표, 다국어 약관)에서 규칙 후보 추출에 LLM 필요. 제어에는 검증 통과 규칙만 사용, LLM 직접 명령 금지',
    '불필요 (로컬 GW에서 정책팩 실행)', '요금이 바뀌어도 재설정 불필요', '신규 요금·프로그램을 수백만 가구에 즉시·정확 반영',
    '요금 API 연동(Tibber, Octopus, gridX)은 흔하나 규제 전반의 검증형 컴파일은 공개 사례 없음', '국가별 검증 규칙 라이브러리와 회귀 시험 세트',
    '요금·계통 규칙을 제약식으로 변환하고 실측 청구 데이터와 대조해 가입 요금을 자동 식별·검증']],
  ['T3', 'Comfort-Bounded Thermal Flex', 'GT-3', 'HP·에어컨·급탕을 쾌적 보장 계약이 붙은 가상 배터리로 — 조절 가능 kW·시간을 신뢰구간과 함께 산출', [
    'DE HP 29.9만 대(80% 기존주택), 사우디 에어컨이 피크의 약 70%, DR override 12.9%, NL 배터리 ROI 약함',
    'HP를 SG-Ready on/off·스케줄로만 다룸. 집마다 단열·재실·기호가 달라 얼마나 줄여도 되는지 모름 → 보수적 제어 또는 override',
    '가정별 열 응답·쾌적 한계 학습 → “N시간 동안 X kW 증감 가능”을 신뢰구간으로 산출 → 한계 안에서만 예열·예냉·급탕 축열, 쾌적 이탈 시 즉시 복귀. 중동은 습도(잠열) 포함',
    '열용량·손실 파라미터와 개인 쾌적 선호를 설치 후 운전 데이터만으로 추정. 쾌적 선호는 모델링 불가 → 학습 필요',
    '부분 필요. 실내 온습도 센서, HP·AC 내부 상태(압축기 주파수, 제상) 접근 — 자사 기기의 강점',
    '배터리 없이도 PV 자가소비·요금 절감, 불편은 계약 범위 내', '최대 가정 부하를 신뢰할 수 있는 flex 자원으로 입찰, 걸프국 냉방 피크 저감',
    'Octopus는 자사 설치 HP 요금 운전, Nest는 온도조절기 DR, 1KOMMA5 AC 진입. 쾌적 보증형 flex 용량 산출은 공개 사례 없음',
    'HP·AC 제조사만 가진 내부 운전 한계 데이터 + 계절 실증 누적(따라잡기 2년+)',
    '(a) 쾌적 이탈 확률을 제약으로 한 flex 용량 산출·입찰 (b) 이벤트 중 재실자 수동 조작을 신호로 쾌적 한계 실시간 갱신']],
  ['T4', 'Grid-Envelope Controller', 'GT-2', '연계점 수입·송출 한도를 법적 수준으로 보장하며 EV·HP·ESS에 1초 이하로 재배분', [
    '§14a DSO 스로틀, 스마트미터 없으면 PV 60% 제한, UK G100(5초), NL 혼잡, IT 3kW 계약, 호주식 DOE 확산',
    '기기별 스로틀만 가능해 연계점 합계를 보장 못함 → DSO는 보수적 한도, 고객은 증설 요구',
    '연계점 측정 → 동적 한도(DSO 신호) → 기기별 비용·쾌적 손실 최소 재배분(<1초) → 감사 가능한 준법 로그',
    '한도 축소 시 어떤 부하를 얼마나 줄일지는 T3 쾌적 envelope·EV 출발 시각·ESS 가치에 달림 → 학습된 가치 함수',
    '핵심. 계량기급 측정, 빠른 제어 출력, 인증 경로', '증설 없이 HP·EV 추가, 빠른 연계 승인, §14a 할인',
    '한도를 신뢰할 수 있어 flexible connection을 더 많이 허가 (EU EMD)', 'G100 장치(SolarEdge, GivEnergy), §14a HEMS — 대부분 PV·ESS 한도 제어에 머묾',
    '자사 HP 압축기 연속 가변 제어 — on/off보다 손실 작은 조정', '동적 연계 한도 하에서 열부하 쾌적 여유와 EV 이동 여유를 공통 단위로 환산해 배분']],
  ['T5', 'Virtual Service Upgrade Hub', 'GT-2', '패널 교체 없이 붙이는 retrofit 허브·미터칼라로 증설 없는 전기화', [
    '미국 단독주택 29.5%가 100A 이하, 서비스 증설 수천~1만 달러+·수개월 대기', 'SW만으로는 NEC 부하 계산 대체 불가 → 인증된 PCS 필요',
    'retrofit 허브·미터칼라에 회로 단위 릴레이·계측 + UL 3141 PCS 로직, ESS·EV·HP 연결 일원화', '차단 없이 한도를 지키려면 부하 예측 기반 선제 배분 필요(T4 공유)',
    '핵심', '전기화 비용 대폭 절감 (업계 추정 $2.5–5k)', '변압기·인입선 증설 이연', 'SPAN(Eaton 투자), Schneider Pulse, ConnectDER, Tesla Backup Switch — 경쟁 치열',
    '미국 패널 자체보다 Energy Pack 일체형 허브 + 유럽 3상 버전', '설치 시 T1 식별 결과로 PCS 한도 설정을 자동 산출·검증']],
  ['T6', 'Resilience Autopilot', '지역 모듈', '정전 위험에 따라 예비 잔량을 조정하고 정전 중 부하 단위 생존시간을 극대화', [
    'Uri 450만 가구 최장 4일, Beryl 폭염 속 수일 정전, 이베리아 대정전, 캘리포니아 PSPS', '“폭풍 예보 시 100% 충전”에서 끝. 정전 중에는 고객이 직접 부하를 끄거나 고정 백업 회로만',
    '정전 위험(기상·계통 경보)별 VPP 방전 대비 예비 잔량 조정 → 정전 중 중요부하 우선순위·PV 충전 예상으로 부하 사이클링(냉장고 듀티, AC 간헐) → 생존시간 신뢰구간 표시',
    '정전 확률·지속시간 추정, 다수 부하 사이클링 스케줄 최적화', '필요. 회로 차단 릴레이(T5), grid-forming 인버터', '같은 배터리로 더 오래, 폭염·한파 시 안전',
    '평시 VPP 참여를 더 허용해도 고객 백업 보장', 'Tesla Storm Watch, SPAN 회로 차단', '자사 AC·냉장고 저전력 운전 모드와 결합', '정전 위험도에 따라 VPP 약정 잔량과 백업 예비량을 확률적으로 배분']],
  ['T7', 'Control-Rights Arbiter', 'GT-4', '누가·언제·어디까지 제어하는지를 계약 우선순위로 기록·중재하는 로컬 권한 원장', [
    'VPP 등록 <1/5·OEM 사일로, TX 소매사·ERCOT·TDU 신호 병존, DE §14a 스로틀 + 공급사 동적요금 동시 수신',
    '대부분 “단일 제어 주체”를 가정. 충돌 시 먼저 온 신호가 이기거나 참여 자체 차단',
    '기기별 제어 권한 원장. 우선순위: 규제 명령(DSO) > 고객 안전·쾌적 > 약정 VPP > 요금 최적화. 모든 결정을 감사 로그로 남기고 정산(T8)과 연결',
    '동시 요청 간 기회비용 비교·배분 결정', '로컬 실행 필요 (클라우드 장애 시에도 DSO 명령 준수)', '여러 프로그램 동시 수익, 공급사 변경 시 재설정 불필요',
    '이미 설치된 DER의 VPP 등록 확대', '공개된 상용 제품 없음. Kraken·Tesla는 자사가 단일 제어자', 'DSO·소매사와의 초기 표준 계약 — 선점 효과 큼',
    '다수 제어 주체의 명령을 계약 우선순위·기회비용으로 중재하고 결과를 정산 가능한 원장에 기록']],
  ['T8', 'Verified Savings & Settlement', 'GT-4', '반사실 기준선으로 실제 절감을 검증하고 성과형 과금·절감 보증·금융으로 연결', [
    'DE 동적요금 역전, UK TOU 전환 의향 17%, NL 회수기간 > 보증, 미국 TPO 전환', '대시보드는 “예상 절감”만. 날씨·재실이 바뀌어 실제 절감 검증 불가',
    '“제어가 없었다면”의 사용량·요금을 추정하는 반사실 기준선 + 무작위 제어 휴지 구간 자체 검증 + 감사 가능 정산 리포트 → 절감 보증·성과형 과금·금융 상품',
    '가정별 반사실 추정은 인과추론 문제', '계량 데이터(GW 측정)', '“효과 없으면 돈을 내지 않는다” — 구매 위험 제거', 'DR·VPP 성과 측정 비용 절감, 분쟁 감소',
    'VPP 기준선 방법론은 있으나 가정 단위 통합 절감 검증은 드묾', '대조군 데이터 누적', '제어 알고리즘이 짧은 무작위 휴지 구간을 스스로 삽입해 실시간 인과 효과를 추정·정산']],
  ['T9', 'Legacy Appliance Enablement', '', 'API 없는 split AC·온수기·구형 HVAC를 NILM과 retrofit 액추에이터로 제어', [
    'UAE·사우디 빌라 다수 split AC, 프랑스 전기온수기 타이머, 미국 구형 HVAC', 'API가 없으면 제어 대상에서 제외',
    '분전반 계측 NILM으로 기기 식별 + IR·릴레이 retrofit 액추에이터 + 계측으로 실제 동작 확인하는 폐루프', '부하 분리, IR 코드·능력 추정',
    '저가 액추에이터 필요', '교체 없이 냉방비·온수비 절감', '기존 설비를 DR 자원화', 'Sensibo 등 IR 컨트롤러, 스마트플러그 — 계통 연계 약함',
    '자사 신형 AC 교체 경로 확보(판매 연결)', 'IR 명령 전후 전력 서명으로 명령 성공·기기 능력 자동 확인']],
  ['T10', 'Mobility-Aware V2H', '', '제조사 중립 V2H와 이동 필요량 보증으로 EV를 가정 배터리로', [
    '독일 2026년부터 V2G 이중 망요금 폐지, EV 배터리는 가정 ESS의 5–8배', 'V2H가 OEM·충전기 조합별로만 가능, 이동 필요량 보장 불명확 → 고객 거부',
    'ISO 15118-20 기반 제조사 중립 V2H, ESS DC 버스 공유, 출발 시각·주행 필요량 보증 제약', '출발 시각·주행거리 예측', 'DC 양방향 모듈 필요',
    '별도 ESS 구매 없이 백업·요금 절감', '대용량 flex 자원', 'Sigenergy DC 25kW 양방향, GM·Ford·Tesla 자사 조합', '표준 차량 보급이 2028년 이후 → 중기 과제',
    '이동 필요량의 확률적 보증을 제약으로 한 V2H 방전 한도 산출']],
  ['T11', 'Remote Performance Assurance', '', '운전 데이터로 원인을 진단하고 원격 자동 재튜닝 후 효과까지 검증', [
    '설치 품질 낮은 HP는 SCOP 15–30%↓(업계), 영국 설치자 부족으로 재방문 비용 큼', '모니터링은 “이상 알림”까지, 원인 파악·교정은 방문 기사',
    '운전 데이터로 원인(히트커브, 유량, 과대 용량) 진단 → 안전 범위 내 원격 자동 재튜닝 → 효과 검증(T8)', '이상 원인 분류', '기존 GW로 충분',
    '운전비 절감, 불편 감소', '겨울 피크 감소', '제조사 원격 진단 서비스 일부', '자사 기기 한정 쉽게 구현, 판매 품질 경쟁력', '원격 재튜닝 전후 자동 A/B 검증·롤백 절차']],
  ['T12', 'Renter Micro Pack', '', '설치공사 없는 800VA 플러그인 PV + 소형 ESS + 자사 가전 협조 운전', [
    '영국 2026.8.27 800VA 플러그인 PV 합법화·저소득층 £25m 시범, 임차인은 설치공사 불가', 'HEMS는 자가 단독주택 설치를 전제',
    '플러그인 PV + 소형 ESS + 자사 가전(세탁기·냉장고·AC) 협조 운전으로 역송 없이 자가소비 극대화', '분전반 계측 없이 소켓 단위로 가정 전체 부하 추정',
    '필요 (플러그인 ESS)', '공사 없이, 이사 시 가져감', '역송 없는 소규모 분산 자원', 'EcoFlow, Anker, Zendure — 가격 경쟁 심함',
    '가전 협조 제어만 차별화 가능, HW 단독 격차 어려움', '소켓 계측·가전 상태로 계량점 전체 역송 추정·방지']]];
details.forEach(d => {
  const s = slide('기술별 상세 · 10개 질문', d[0] + '  ' + d[1], '', '특허 아이디어는 선행기술 조사 전의 후보. 수치 출처는 국가별 슬라이드 노트 및 원문 보고서 참조.');
  card(s, 0.6, 1.5, 3.0, 5.35, C.panel);
  T(s, d[0], 0.85, 1.7, 2.5, 0.7, { fontSize: 36, bold: true, color: C.blue });
  if (d[2]) pill(s, d[2], 0.85, 2.5, d[2].startsWith('GT') ? 'gt' : 'mid');
  T(s, d[3], 0.85, 3.0, 2.55, 2.4, { fontSize: 12.5 });
  const r = techRows.find(x => x[0].startsWith(d[0] + ' '));
  T(s, [{ text: 'Target  ', options: { bold: true, color: C.muted } }, { text: r[2], options: { breakLine: true } },
    { text: '차별화  ', options: { bold: true, color: C.muted } }, { text: r[8].text, options: { breakLine: true } },
    { text: '필요성  ', options: { bold: true, color: C.muted } }, { text: r[9].text }], 0.85, 5.55, 2.55, 1.2, { fontSize: 10.5, paraSpaceAfter: 3 });
  const rows = Q.map((q, i) => [{ text: q, options: { bold: true, color: C.muted, fill: { color: i % 2 ? C.white : C.panel } } }, { text: d[4][i], options: { fill: { color: i % 2 ? C.white : C.panel } } }]);
  s.addTable(rows, { x: 3.8, y: 1.5, w: 8.93, colW: [1.95, 6.98], fontFace: F, fontSize: 11.5, rowH: 0.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [3, 6, 3, 6], autoPage: false });
});

// ---------- Chain ----------
{
  const s = slide('Gateway Technology와 가치사슬', 'Country Pain Point → Business Opportunity', '기술에서 출발하지 않고 Pain Point에서 거꾸로 추적 — 각 행은 “이 기술이 없어서 지금 시장이 못 크고 있다”는 연결');
  const H = ['Country Pain Point', 'Market Barrier', 'Unsolved Problem', 'Gateway Technology', 'Product / Energy Pack', 'Customer Value', 'Business Opportunity'];
  const R = [
    ['DE·UK·US\n설치자 부족, 통합 2–8k€, soft cost 58%', '설치 1건당 비용·시간이 가치보다 큼', '혼합 제조사 기기를 사람 없이 정확히 연결·검증', 'GT-1 Zero-Touch Commissioning (T1+T2)', 'Energy Pack 게이트웨이 + 설치 앱 + 국가 정책팩', '설치 당일 정상 작동, 설치비 절감', '설치업체·OEM 라이선스, 타사 HP·ESS 설치 채널'],
    ['US·DE·UK·NL·IT\n100A 패널, §14a·60%, G100, 혼잡', '연계 용량이 전기화 속도 제한', '연계점 한도를 법적으로 보장하며 HP·EV·ESS 동시 사용', 'GT-2 Grid-Envelope Power Hub (T4+T5)', 'Power Hub (계량기급 측정 + 회로 제어 + PCS)', '증설 없는 전기화, 빠른 연계 승인', 'HW 판매 + DSO·utility 조달(flex connection)'],
    ['DE·NL·UK·TX·GCC\nHP·AC가 최대 부하, 배터리 ROI 약함, override', '최대 유연 자원이 시장에 못 들어옴', '쾌적을 해치지 않는 가정별 flex 용량 보증', 'GT-3 Comfort-Bounded Thermal Flex (T3)', '자사 HP·AC + Energy Pack(배터리 선택)', '배터리 없이 절감, 불편 없음', 'HP·AC 판매 경쟁력, flex 용량 판매, 걸프국 B2G'],
    ['US·CA·TX·UK·DE\nROI 불확실, VPP 등록 <1/5, lock-in', '수익이 쪼개지고 증명 안 돼 구매·금융 막힘', '다수 제어 주체 중재 + 검증된 가치 정산', 'GT-4 Multi-Party Control & Verified Value (T7+T8)', '로컬 권한 중재 + 정산 서비스', '다중 프로그램 동시 수익, 성과 보증', '성과형 구독, 절감 보증, 금융·보험 제휴'],
    ['TX·CA·ES\n폭염·한파·대정전', '백업 신뢰 부족', '부하 단위 생존시간 극대화', '지역 모듈: Resilience Autopilot (T6)', 'Power Hub + grid-forming ESS', '정전 중 안전', '프리미엄 백업 번들, 소매사 리스(TX)']];
  const body = [H.map(hdr), ...R.map((r, i) => r.map((c, j) => ({ text: c, options: { bold: j === 0 || j === 3, color: j === 3 ? (i < 4 ? C.copper : C.muted) : C.ink, fill: { color: j === 3 ? C.copperSoft : (i % 2 ? C.panel : C.white) } } })))];
  s.addTable(body, { x: 0.6, y: 1.85, w: 12.13, colW: [1.95, 1.6, 1.8, 1.85, 1.7, 1.45, 1.78], fontFace: F, fontSize: 12.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [4, 5, 4, 5], autoPage: false });
}

// ---------- Roadmap ----------
{
  const s = slide('실행 제안', '우선순위와 이 전략이 틀릴 수 있는 조건', '');
  const steps = [['2027년 착수', 'GT-1 + GT-3', '공통 게이트웨이 SW, HP·AC 내부 데이터 활용, 독일·영국 실증'],
    ['2027–2028년', 'GT-2 · GT-4', '유럽 3상 허브(§14a·G100), 권한 원장을 DSO·소매사 1곳과 공동 설계'],
    ['2028년 이후', '확장', '미국 PCS 허브·T6(TX·CA 파트너), T10 V2H, 걸프국 냉방 flex B2G']];
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.35, w: 7.2, h: 0, line: { color: C.line, width: 2 } });
  steps.forEach((a, i) => {
    const x = 0.6 + i * 2.55;
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.2, w: 0.3, h: 0.3, fill: { color: C.copper }, line: { color: C.white, width: 2 } });
    T(s, a[0], x, 2.7, 2.4, 0.35, { fontSize: 14, bold: true, color: C.copper });
    T(s, a[1], x, 3.1, 2.4, 0.4, { fontSize: 17, bold: true, color: C.blue });
    T(s, a[2], x, 3.6, 2.3, 1.6, { fontSize: 12 });
  });
  card(s, 8.4, 1.9, 4.33, 4.9, C.panel);
  T(s, '틀릴 수 있는 조건', 8.7, 2.1, 3.8, 0.4, { fontSize: 15, bold: true, color: C.bad });
  T(s, ['Matter·EEBus 표준화가 빨라져 커미셔닝 난이도가 크게 낮아지는 경우 (GT-1 가치 감소)', '1KOMMA5·Octopus가 제조사 중립 레트로핏을 대규모 확산', '자사 HP·AC 내부 상태 API 접근을 사내에서 확보하지 못하는 경우 (GT-3 격차 소멸)', '걸프국 utility가 가정 DR 예산을 배정하지 않는 경우']
    .map((t, k, a) => ({ text: t, options: { bullet: { indent: 15 }, breakLine: k < a.length - 1 } })), 8.7, 2.6, 3.8, 4, { fontSize: 12, paraSpaceAfter: 9 });
}

// ---------- Closing ----------
{
  const s = slide('한 줄 요약', '', '', '', true);
  T(s, '2027–2030 Energy Pack 경쟁의 승부처는\n최적화 알고리즘이 아니라 통합 능력이다', 0.8, 1.5, 11.7, 1.5, { fontSize: 30, bold: true, color: C.white });
  const L = [['GT-1', '연결하는 순간 알아서 맞춰지고'], ['GT-2', '계통 한도 안에서'], ['GT-3', '냉난방까지 불편 없이'], ['GT-4', '누구의 명령이든 중재해 검증된 돈으로 바꾼다']];
  L.forEach((a, i) => {
    const y = 3.35 + i * 0.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y, w: 0.95, h: 0.42, fill: { color: C.copper }, line: { color: C.copper }, rectRadius: 0.21 });
    T(s, a[0], 0.8, y, 0.95, 0.42, { fontSize: 13, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T(s, a[1], 2.0, y, 10, 0.42, { fontSize: 18, color: C.white, valign: 'middle' });
  });
  T(s, '냉난방 기기와 Energy Pack을 함께 가진 회사만 네 가지를 한 번에 갖출 수 있다.', 0.8, 6.05, 11.7, 0.5, { fontSize: 16, bold: true, color: 'E0A06A' });
}

// ---------- Sources ----------
const SRC = [
  'BNetzA 동적요금 · bundesnetzagentur.de/DE/Vportal/Energie/Vertragsarten/DynStromtarife', 'BNetzA §14a · bundesnetzagentur.de/DE/Vportal/Energie/SteuerbareVBE', 'Utility Partners 스마트미터·§14a † · utility-partners.de',
  'Verivox 동적요금 비교(blackout-news) † · blackout-news.de', 'pv magazine Solarspitzengesetz 2025.2 · pv-magazine.com/2025/02/17', 'pv magazine 독일 배터리 2025 · pv-magazine.com/2026/01/12 · pv-magazine.de/2026/01/06',
  'BWP 독일 HP 2025 · renewable-energy-industry.com article-7202', 'Mobility House V2G 법 개정 · mobilityhouse.com', 'heatload HEMS Integration 2026 † · heatload.ing', 'Vattenfall §14a Modul · vattenfall.de',
  'Eurostat 2025 가정 전기요금 · ec.europa.eu/eurostat ddn-20260505-1', 'Ofgem price cap · ofgem.gov.uk', 'Ofgem 2025 flexibility consumer survey · ofgem.gov.uk/research', 'GOV.UK HP deployment stats June 2026 · Envirotec −22%',
  'MCS Foundation HP skills · mcsfoundation.org.uk', 'Elexon MHHS · elexon.co.uk', 'NESO Demand Flexibility Service · neso.energy', 'Fuse Energy 플러그인 태양광 † · fuseenergy.com',
  'EV Infrastructure News Kraken 2GW', 'Reonic G100 † · reonic.com', 'Enedis Heures creuses 2025.10.15 · enedis.fr', 'photovoltaique.info 매입 요금', 'Selectra Linky † · selectra.info', 'CAPEB MaPrimeRénov · capeb.fr',
  'Rijksoverheid Salderingsregeling · rijksoverheid.nl', 'ENTSO-E Iberian blackout final report', 'The Local Italy 2026.9', 'SolarPower Europe EU battery 2025', 'European Commission EMD', 'IEA Electricity 2026 Flexibility',
  'CPUC Net Billing Tariff · cpuc.ca.gov', 'Solar.com PG&E rates 2026 †', 'Brattle VPP test 2025.8', 'pv magazine USA DSGS 삭감', 'SEIA OBBB clean energy provisions', 'ACEEE 2024 panel capacity survey',
  'Solar Builder soft costs †', 'US DOE VPP Liftoff 2025 Update · liftoff.energy.gov', 'Utility Dive / EIA 데이터센터·가정요금', 'Utility Dive Tesla VPP ERCOT', 'Texas Tribune Hurricane Beryl',
  'Electrek Tesla Powerwall 리스 2026.8', 'Texas Energy Fund TBPP · txenergyfund.texas.gov', 'DEWA slab tariff · dewa.gov.ae', 'SurgePV Shams Dubai †', 'Dubai SCE DSM Strategy 2050 · dubaisce.gov.ae',
  'SERA consumption tariff · sera.gov.sa', 'ScienceDirect 2025 사우디 PV 인센티브 · vision2030.ai †', 'Kahramaa BeSolar · km.qa', 'ESS News 1KOMMA5 1GW · 1komma5.com Heartbeat',
  'Tesla Storm Watch · tesla.com', 'Canary Media SPAN Edge', 'Enphase IQ Energy Management 2026.3', 'pv Europe Huawei Smart Assistant', 'Samsung SmartThings Energy', 'Schneider Wiser Home AI',
  'sonnen Europe largest VPP', 'CSA Matter 1.4 · csa-iot.org', 'ScienceDirect 2025 smart thermostat review', 'POWER Voltus–Octopus', 'Sigenergy SigenStor', 'Applied Energy 2026 DOE review',
  'Mitsubishi Electric §14a', 'EcoFlow OCEAN Pro', 'gridX Salderingsregeling †', 'CPUC Rule 21', 'Carbon Brief Warm Homes Plan', 'TPPF Uri 이후 계통 취약성'];
[[0, 34], [34, SRC.length]].forEach(([a, b], k) => {
  const s = slide('출처', '참고 자료 (' + (k + 1) + '/2)', '2025–2026 자료 우선 · † 업계·2차 자료(방향성 근거) · 전체 URL은 원문 HTML 보고서 9장에 수록');
  const items = SRC.slice(a, b).map((t, i) => (a + i + 1) + '. ' + t);
  const half = Math.ceil(items.length / 2);
  [items.slice(0, half), items.slice(half)].forEach((col, j) => T(s, col.join('\n'), 0.6 + j * 6.13, 1.9, 5.93, 5.0, { fontSize: 9.5, color: C.ink, lineSpacingMultiple: 1.15 }));
});

pres.writeFile({ fileName: OUT }).then(f => console.log('written', f));
