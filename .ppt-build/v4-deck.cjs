// HEMS·Energy Pack v4 deck — content parsed from output/Energy_Pack_Gateway_Technology_v4_Full.html
const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');
const SRC_HTML = path.resolve(__dirname, '../output/Energy_Pack_Gateway_Technology_v4_Full.html');
const OUT = process.argv[2] || path.resolve(__dirname, '../output/Energy_Pack_Gateway_Technology_v4_Full.pptx');
const html = fs.readFileSync(SRC_HTML, 'utf8');

// ---------- HTML helpers ----------
const ent = s => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&quot;/g, '"');
const noRef = s => s.replace(/<a class="r"[^>]*>.*?<\/a>/g, '');
const plain = s => ent(noRef(s).replace(/<br\s*\/?>/g, '\n').replace(/<[^>]+>/g, '')).replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').trim();
function runs(s, base = {}) { // rich text runs with <b> and <br>
  s = noRef(s);
  const out = []; const re = /<b>([\s\S]*?)<\/b>|<br\s*\/?>|<[^>]+>|[^<]+/g; let m;
  while ((m = re.exec(s))) {
    if (m[1] !== undefined) out.push({ text: ent(m[1].replace(/<[^>]+>/g, '')), options: Object.assign({}, base, { bold: true }) });
    else if (/^<br/.test(m[0])) out.push({ text: '', options: Object.assign({}, base, { breakLine: true }) });
    else if (m[0][0] === '<') continue;
    else out.push({ text: ent(m[0]).replace(/\s+/g, ' '), options: Object.assign({}, base) });
  }
  // merge breakLine markers into previous run
  const res = [];
  out.forEach(r => { if (r.text === '' && r.options.breakLine) { if (res.length) res[res.length - 1].options.breakLine = true; } else res.push(r); });
  if (res.length) { res[0].text = res[0].text.replace(/^\s+/, ''); delete res[res.length - 1].options.breakLine; }
  return res;
}
const section = id => { const m = html.match(new RegExp(`<section id="${id}">([\\s\\S]*?)</section>`)); return m ? m[1] : ''; };
const tables = seg => [...seg.matchAll(/<table[^>]*>([\s\S]*?)<\/table>/g)].map(t => [...t[1].matchAll(/<tr>([\s\S]*?)<\/tr>/g)].map(r => [...r[1].matchAll(/<t([dh])([^>]*)>([\s\S]*?)<\/t[dh]>/g)].map(c => ({ tag: c[1], attr: c[2], html: c[3] }))));
const refs = seg => [...new Set([...seg.matchAll(/href="#src(\d+)"/g)].map(m => +m[1]))].sort((a, b) => a - b);
const SOURCES = {}; [...html.matchAll(/<li id="src(\d+)">([\s\S]*?)<\/li>/g)].forEach(m => { SOURCES[+m[1]] = plain(m[2]); });
const notesFor = (seg, extra = '') => { const r = refs(seg); return (extra ? extra + '\n\n' : '') + (r.length ? '출처\n' + r.map(n => `[${n}] ${SOURCES[n]}`).join('\n') : ''); };

// ---------- design ----------
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';
pres.title = 'HEMS와 Energy Pack 2027–2028 신기술';
const F = 'Malgun Gothic';
const C = { dark: '142233', dark2: '1F3048', ink: '18212B', blue: '1D4F91', blueSoft: 'E3EBF6', copper: 'B8641F', copperSoft: 'F6EADF',
  muted: '5B6773', line: 'D3DAE1', panel: 'F2F4F7', white: 'FFFFFF', ok: '2D6A4F', okSoft: 'E1EFE7', warn: '8A5A00', warnSoft: 'F7ECD2',
  bad: '9B2C2C', badSoft: 'F5E1E1', darkMuted: 'A9B8C8', lightCopper: 'E0A06A' };
let n = 0;
function T(s, text, x, y, w, h, o = {}) { s.addText(text, Object.assign({ x, y, w, h, fontFace: F, fontSize: 14, color: C.ink, margin: 0, valign: 'top', isTextBox: true }, o)); }
function R(s, x, y, w, h, fill, o = {}) { s.addShape(o.round ? pres.shapes.ROUNDED_RECTANGLE : pres.shapes.RECTANGLE, Object.assign({ x, y, w, h, fill: { color: fill }, line: { color: o.lineColor || fill, width: o.lineW || 0.75 } }, o.round ? { rectRadius: o.round } : {})); }
function slide(sec, title, sub, notes, dark = false) {
  n++; const s = pres.addSlide(); s.background = { color: dark ? C.dark : C.white };
  if (sec) T(s, sec, 0.6, 0.36, 10, 0.28, { fontSize: 11, color: C.copper, bold: true, charSpacing: 1 });
  if (title) T(s, title, 0.6, 0.64, 12.1, 0.62, { fontSize: 26, bold: true, color: dark ? C.white : C.ink, fit: 'shrink' });
  if (sub) T(s, sub, 0.6, 1.27, 12.1, 0.38, { fontSize: 13, color: dark ? C.darkMuted : C.muted });
  T(s, 'HEMS · Energy Pack 신기술 v4  ·  2026.09.30', 0.6, 7.08, 8, 0.25, { fontSize: 9, color: dark ? C.darkMuted : C.muted });
  T(s, String(n).padStart(2, '0'), 12.1, 7.08, 0.63, 0.25, { fontSize: 9, color: dark ? C.darkMuted : C.muted, align: 'right' });
  if (notes) s.addNotes(notes);
  return s;
}
function pill(s, text, x, y, kind, h = 0.3, fs = 10) {
  const m = { hi: [C.okSoft, C.ok], mid: [C.warnSoft, C.warn], lo: [C.badSoft, C.bad], gt: [C.copperSoft, C.copper], blue: [C.blueSoft, C.blue], dark: ['22344A', C.white] }[kind];
  const w = Math.max(0.6, 0.105 * [...text].reduce((a, ch) => a + (/[ㄱ-힝]/.test(ch) ? 1.55 : 0.85), 0) + 0.32);
  R(s, x, y, w, h, m[0], { round: h / 2 });
  T(s, text, x, y, w, h, { fontSize: fs, bold: true, color: m[1], align: 'center', valign: 'middle' });
  return w;
}
function bullets(items, o = {}) { return items.map((t, k) => ({ text: t, options: Object.assign({ bullet: { indent: 14 }, breakLine: k < items.length - 1 }, o) })); }
const HDR = t => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.dark } } });
function tableSlideFromRows(s, rows, colW, o = {}) {
  const body = rows.map((r, i) => r.map((c, j) => {
    if (i === 0) return HDR(typeof c === 'string' ? c : plain(c.html));
    if (typeof c === 'string') return { text: c, options: { bold: j === 0, fill: { color: i % 2 ? C.white : C.panel } } };
    const txt = runs(c.html);
    return { text: txt.length ? txt : ' ', options: { bold: j === 0 && o.boldFirst !== false, fill: { color: i % 2 ? C.white : C.panel } } };
  }));
  s.addTable(body, Object.assign({ x: 0.6, y: o.y || 1.8, w: colW.reduce((a, b) => a + b, 0), colW, fontFace: F, fontSize: o.fs || 11, color: C.ink,
    border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'top', margin: [4, 6, 4, 6], autoPage: false }, o.opts || {}));
}
function divider(num, title, sub) {
  const s = slide('', '', '', '', true);
  T(s, num, 0.8, 2.2, 3, 0.5, { fontSize: 16, bold: true, color: C.copper, charSpacing: 2 });
  T(s, title, 0.8, 2.75, 11.5, 1.0, { fontSize: 38, bold: true, color: C.white });
  if (sub) T(s, sub, 0.8, 3.85, 11.5, 0.9, { fontSize: 16, color: C.darkMuted });
  return s;
}

// ================= 1. Title =================
{
  const s = slide('', '', '', '원문: output/Energy_Pack_Gateway_Technology_v4_Full.html (근거 자료 81건)', true);
  T(s, 'HEMS · ENERGY PACK 신기술 · FULL v4', 0.8, 1.45, 11, 0.35, { fontSize: 13, color: C.copper, bold: true, charSpacing: 2 });
  T(s, 'HEMS와 Energy Pack\n2027–2028 신기술', 0.8, 1.95, 11.5, 1.9, { fontSize: 44, bold: true, color: C.white });
  T(s, 'HEMS는 모든 집의 두뇌, Energy Pack은 그 두뇌가 가장 큰 힘을 내는 HW 패키지', 0.8, 4.0, 11.5, 0.5, { fontSize: 18, color: C.darkMuted });
  let x = 0.8; ['유럽 · 독일 영국 프랑스 네덜란드 남유럽', '북미 · California Texas', '중동 · UAE 사우디 카타르', 'PL 포함 6명'].forEach(t => { x += pill(s, t, x, 4.95, 'dark', 0.42, 11.5) + 0.2; });
  T(s, '2026.09.30 · 2025–2026 공식·규제기관·업계 자료 81건', 0.8, 5.75, 11, 0.3, { fontSize: 11, color: C.darkMuted });
}

// ================= 2. 결론 =================
const s0 = section('s0');
{
  const s = slide('결론', '내년에 이 과제로 무엇을 하는가', '', notesFor(s0));
  R(s, 0.6, 1.5, 12.13, 1.55, C.copperSoft);
  T(s, [{ text: 'Energy Pack이 있는 집이든 우리 가전만 있는 집이든, ', options: {} },
    { text: '연결하는 순간 스스로 설정되고(GT-1), 냉난방을 불편 없이 조절하고(GT-2), 가전 안에서 바로 동작하며(GT-3), 절감을 돈으로 증명하는(GT-4)', options: { color: C.copper } },
    { text: ' HEMS를 만들고, Energy Pack에는 계통 한도를 지키는 Power Hub(GT-5)를 더한다.', options: {} }], 0.9, 1.62, 11.55, 1.3, { fontSize: 17, bold: true, valign: 'middle' });
  const cols = [
    ['HEMS', C.blue, C.blueSoft, '집의 에너지 두뇌(SW·AI). Energy Pack이 없어도 자사 에어컨·히트펌프·세탁기·건조기·온수기만으로 동작.', 'GT-1 · GT-2 · GT-3 · GT-4'],
    ['Energy Pack', C.copper, C.copperSoft, 'PV·ESS·히트펌프·에어컨을 묶은 HW 패키지. HEMS가 가장 큰 가치를 내는 집이지만 유일한 무대는 아님.', 'GT-5 Power Hub'],
    ['진입 경로', C.ink, C.panel, 'A Energy Pack 설치 집 · B 자사 가전만 있는 집 · C 타사 기기가 섞인 집. Energy Pack이 늦게 팔려도 B·C로 HEMS 고객이 쌓임.', 'A · B · C']];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 4.1; R(s, x, 3.3, 3.9, 3.45, c[2]);
    T(s, c[0], x + 0.3, 3.5, 3.3, 0.45, { fontSize: 20, bold: true, color: c[1] });
    T(s, c[3], x + 0.3, 4.05, 3.3, 1.95, { fontSize: 15 });
    T(s, c[4], x + 0.3, 6.15, 3.3, 0.4, { fontSize: 12, bold: true, color: c[1] });
  });
}
{
  const s = slide('결론', 'Gateway Technology 5개', 'HEMS 4개 + Energy Pack HW 1개 · 예측 정확도·대시보드·Digital Twin·단독 Autopilot은 제외', notesFor(s0));
  const g = [['GT-1', 'HEMS 연결', 'Zero-Touch Commissioning', '타사 기기까지 자동 발견하고 시험 신호로 배선·제어권 확인. 설치기사 수작업 2–4시간 → 앱 확인 수준.'],
    ['GT-2', 'HEMS 제어', 'Comfort-Bounded Thermal Flex', '싼 시간에 집에 열을 미리 저장하고 비싼 시간엔 덜 돌려 배터리 없이 절감. 사용자 설정 변경을 선호로 학습.'],
    ['GT-3', 'HEMS 확산', 'Appliance-Embedded HEMS & Fleet Flex', '게이트웨이 없이 AC·HP 안에서 HEMS가 동작. 이미 팔린 가전을 묶어 가전 VPP로.'],
    ['GT-4', 'HEMS 가치', 'Verified Value & Control Arbitration', '“제어 없었다면 £126, 실제 £89” 검증 절감, 요금제 자동 선택, 소매사·VPP·DSO 명령 중재.'],
    ['GT-5', 'Energy Pack HW', 'Grid-Envelope Power Hub', '계측·회로 제어 HW로 §14a·G100·100A 패널 한도를 1초 안에 지키며 HP·EV·ESS 동시 운전.'],
    ['제외', '신기술 아님', '연장선 기술은 Enabler로', '예측 정확도 향상, 대시보드, Digital Twin, 단독 Energy Autopilot. 기존 예측·최적화·HVAC 제어는 5개의 기반.']];
  g.forEach((a, i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.85 + Math.floor(i / 3) * 2.6; const hw = a[0] === 'GT-5', ex = a[0] === '제외';
    R(s, x, y, 3.9, 2.4, ex ? C.panel : hw ? C.copperSoft : C.blueSoft);
    const w = pill(s, a[0], x + 0.25, y + 0.22, ex ? 'mid' : 'gt');
    T(s, a[1], x + 0.35 + w, y + 0.22, 2.5, 0.3, { fontSize: 11, color: C.muted, valign: 'middle' });
    T(s, a[2], x + 0.25, y + 0.65, 3.45, 0.62, { fontSize: 15.5, bold: true, color: ex ? C.muted : hw ? C.copper : C.blue, valign: 'middle' });
    T(s, a[3], x + 0.25, y + 1.32, 3.45, 1.0, { fontSize: 12.5 });
  });
}

// ================= Part 1 divider =================
divider('PART 1', '무엇을 만드는가', 'HEMS·Energy Pack 구조 · Gateway Technology 5개 · 6명 실행계획');

// ================= 구조 =================
const s1 = section('s1');
{
  const s = slide('1 · HEMS와 Energy Pack의 구조', '두뇌와 몸을 나눠서 본다', '', notesFor(s1));
  const layers = [['HEMS 신기술', 'SW·AI · 모든 집', C.blue, C.blueSoft, ['GT-1 Zero-Touch Commissioning', 'GT-2 Thermal Flex', 'GT-3 가전 내장 HEMS·가전 VPP', 'GT-4 검증 절감·요금제·권한 중재', 'T6 정전 대응(지역)'], 4],
    ['Energy Pack 신기술', 'HW · Energy Pack 집', C.copper, C.copperSoft, ['GT-5 Grid-Envelope Power Hub', 'Energy Pack Gateway', 'T13 일체형 HP/AC', 'T10 V2H', 'T12 임차인 Micro Pack'], 1],
    ['국가 적응', '기술 8–9', C.ink, C.panel, ['요금·규제 적응(전력요금 문서 Agent → 규칙팩)', '기기·프로토콜 연동(EEBus·SunSpec·OCPP·Matter)'], 0],
    ['확보 기술', '기술 1–7 · Enabler', C.muted, C.panel, ['PV 발전 예측', '부하 예측', 'ESS 최적 제어', 'HVAC·가전 제어', 'TOU 최적화', '재실·스케줄링', '실증 절감 알고리즘'], 0]];
  layers.forEach((L, i) => {
    const y = 1.55 + i * 1.33; R(s, 0.6, y, 2.6, 1.18, L[3]); R(s, 3.25, y, 9.48, 1.18, C.white, { lineColor: C.line });
    T(s, L[0], 0.8, y + 0.2, 2.3, 0.4, { fontSize: 15, bold: true, color: L[2] });
    T(s, L[1], 0.8, y + 0.65, 2.3, 0.35, { fontSize: 10.5, color: C.muted });
    let x = 3.45, yy = y + 0.2;
    L[4].forEach((c, k) => {
      const w = 0.105 * [...c].reduce((a, ch) => a + (/[ㄱ-힝]/.test(ch) ? 1.55 : 0.85), 0) + 0.3;
      if (x + w > 12.55) { x = 3.45; yy += 0.44; }
      const isNew = k < L[5]; R(s, x, yy, w, 0.34, isNew ? C.copperSoft : C.panel, { lineColor: isNew ? C.copper : C.line });
      T(s, c, x, yy, w, 0.34, { fontSize: 10.5, bold: isNew, color: isNew ? C.copper : C.ink, align: 'center', valign: 'middle' });
      x += w + 0.12;
    });
  });
}
{
  const t = tables(s1)[0];
  const s = slide('1 · HEMS와 Energy Pack의 구조', '세 가지 고객 진입 경로', 'Energy Pack이 늦게 팔려도 경로 B·C로 HEMS 고객과 데이터가 쌓인다', notesFor(s1));
  tableSlideFromRows(s, t, [1.45, 2.55, 1.9, 1.75, 2.45, 2.03], { fs: 12.5 });
  const cards = [['왜 HEMS를 Energy Pack과 분리하나', ['Energy Pack은 설치가 크고 판매 주기가 김. HEMS가 묶이면 판매량만큼만 큼', '자사 가전은 이미 많은 집에 있음. 경로 B로 먼저 쌓고 경로 A로 전환', '경쟁사도 이 방향: Samsung Flex Connect(가전·TV DR), Nest Rush Hour Rewards 130만 가구·연 $40M']],
    ['표준이 가전 HEMS를 밀어준다', ['Matter 1.4: 가전이 소비 예측 보고, 시작 지연·일시정지·출력 조절', 'Matter 1.5(2025.11): 요금·탄소 정보를 기기에 직접 전달', 'EU 스마트가전 상호운용 규약: Daikin·Mitsubishi·Panasonic 서명, 2026.3 EMS·배터리·충전기로 확대']]];
  cards.forEach((c, i) => {
    const x = 0.6 + i * 6.13; R(s, x, 5.0, 5.93, 1.9, i ? C.blueSoft : C.panel);
    T(s, c[0], x + 0.3, 5.1, 5.4, 0.35, { fontSize: 13.5, bold: true, color: C.blue });
    T(s, bullets(c[1]), x + 0.3, 5.47, 5.4, 1.4, { fontSize: 10.5, paraSpaceAfter: 2 });
  });
}

// ================= GT build slides =================
const s2 = section('s2');
const builds = [...s2.matchAll(/<div class="build( hw)?">([\s\S]*?)(?=<div class="build|<h3>사용자)/g)];
builds.forEach(b => {
  const seg = b[2]; const hw = !!b[1];
  const code = plain(seg.match(/<span class="code">([\s\S]*?)<\/span>/)[1]);
  const title = plain(seg.match(/<h3>([\s\S]*?)<\/h3>/)[1]);
  const desc = plain(seg.match(/<\/h3><p>([\s\S]*?)<\/p>/)[1]);
  const rows = {}; [...seg.matchAll(/<div class="row"><b>(.*?)<\/b><span[^>]*>([\s\S]*?)<\/span><\/div>/g)].forEach(m => rows[m[1]] = m[2]);
  const before = seg.match(/<div class="b"><span class="h">.*?<\/span>([\s\S]*?)<\/div>/)[1];
  const after = seg.match(/<div class="a"><span class="h">.*?<\/span>([\s\S]*?)<\/div>/)[1];
  const kpis = [...(rows['목표'] || '').matchAll(/<span>([\s\S]*?)<\/span>/g)].map(m => plain(m[1]));
  const gt = code.split('·')[0].trim();
  const s = slide('2 · 무엇을 만드나 · ' + gt, title, desc, notesFor(seg, code));
  const accent = hw ? C.copper : C.blue, soft = hw ? C.copperSoft : C.blueSoft;
  // left column
  R(s, 0.6, 1.8, 3.25, 5.05, soft);
  T(s, gt, 0.85, 1.98, 2.8, 0.6, { fontSize: 30, bold: true, color: accent });
  T(s, code.split('·').slice(1).join('·').trim(), 0.85, 2.62, 2.8, 0.6, { fontSize: 11.5, bold: true, color: C.copper });
  T(s, '목표', 0.85, 3.3, 2.8, 0.3, { fontSize: 11, bold: true, color: C.muted });
  T(s, bullets(kpis, { color: C.ink }), 0.85, 3.62, 2.8, 2.0, { fontSize: 12.5, paraSpaceAfter: 5 });
  T(s, '일정', 0.85, 5.7, 2.8, 0.3, { fontSize: 11, bold: true, color: C.muted });
  T(s, plain(rows['일정'] || ''), 0.85, 6.0, 2.8, 0.8, { fontSize: 11.5 });
  // right
  T(s, '만드는 것', 4.1, 1.8, 3, 0.3, { fontSize: 12, bold: true, color: accent });
  T(s, runs(rows['만드는 것'] || ''), 4.1, 2.12, 8.63, 2.3, { fontSize: 13, paraSpaceAfter: 3, valign: 'top' });
  R(s, 4.1, 4.55, 4.23, 2.3, C.badSoft); R(s, 8.5, 4.55, 4.23, 2.3, C.okSoft);
  T(s, '지금', 4.3, 4.68, 3.8, 0.3, { fontSize: 11, bold: true, color: C.bad });
  T(s, gt + ' 이후', 8.7, 4.68, 3.8, 0.3, { fontSize: 11, bold: true, color: C.ok });
  T(s, plain(before), 4.3, 5.02, 3.85, 1.75, { fontSize: 13 });
  T(s, plain(after), 8.7, 5.02, 3.85, 1.75, { fontSize: 13 });
});
{
  const seg = s2.slice(s2.indexOf('<h3>사용자'));
  const items = [...seg.matchAll(/<span class="when">(.*?)<\/span><span class="what">(.*?)<\/span><span>(.*?)<\/span>/g)].map(m => m.slice(1).map(plain));
  const s = slide('2 · 무엇을 만드나', '사용자 입장에서 본 한 달', '같은 HEMS가 Energy Pack 집과 가전만 있는 집에서 어떻게 보이는가', '');
  s.addShape(pres.shapes.LINE, { x: 0.9, y: 2.35, w: 11.5, h: 0, line: { color: C.line, width: 2 } });
  items.forEach((a, i) => {
    const x = 0.6 + i * 3.08;
    s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.2, w: 0.3, h: 0.3, fill: { color: C.copper }, line: { color: C.white, width: 2 } });
    T(s, a[0], x, 2.75, 2.9, 0.4, { fontSize: 18, bold: true, color: C.copper });
    pill(s, a[1], x, 3.25, 'blue', 0.32, 10.5);
    T(s, a[2], x, 3.8, 2.85, 2.8, { fontSize: 16 });
  });
}

// ================= 6명 =================
const s3 = section('s3');
{
  const [roles, gantt] = tables(s3);
  const s = slide('3 · 6명 실행계획', '누가 무엇을 만드나 (PL 1 + 팀원 5)', '', notesFor(s3));
  tableSlideFromRows(s, roles, [0.95, 1.9, 4.2, 1.75, 1.75, 1.58], { fs: 13 });
  const s2_ = slide('3 · 6명 실행계획', '분기별 일정', 'HEMS 과제(파랑)는 2027년에 성과를 내고, Energy Pack HW(구리색)는 2027 하반기 설계부터', '');
  const body = gantt.map((r, i) => r.map((c, j) => {
    if (i === 0) return HDR(plain(c.html));
    const cls = (c.attr.match(/class="([^"]+)"/) || [])[1] || '';
    if (j === 0) return { text: plain(c.html), options: { bold: true, fill: { color: C.white } } };
    const fill = cls === 'q-on' ? C.blueSoft : cls === 'q-hw' ? C.copperSoft : C.white;
    const col = cls === 'q-on' ? C.blue : cls === 'q-hw' ? C.copper : C.line;
    return { text: plain(c.html), options: { fill: { color: fill }, color: col, bold: cls !== 'q-off', align: 'center', valign: 'middle' } };
  }));
  s2_.addTable(body, { x: 0.6, y: 1.9, w: 12.13, colW: [3.33, 1.25, 1.25, 1.3, 1.25, 1.25, 1.25, 1.25], rowH: 0.62, fontFace: F, fontSize: 12, color: C.ink, border: { type: 'solid', pt: 0.75, color: C.line }, valign: 'middle', margin: [4, 6, 4, 6], autoPage: false });
}
{
  const s = slide('3 · 6명 실행계획', '12주 실증, 판단 기준, 하지 않는 것, 특허', '', notesFor(s3));
  const cards = [['12주 탐색 실증 (2027 상반기)', ['1–3주: 독일·영국 고객 10명, 설치기사 5명 인터뷰, 기기 접근 확인', '4–6주: 지원 10기종 설치 시험, 기존 HEMS와 설정 시간 비교', '7–10주: 독일·영국 각 10호 실증(GT-2·4), 자사 AC 내장 HEMS 시험(GT-3)', '11–12주: 계속/축소 판단'], C.blueSoft, C.blue],
    ['계속/축소 판단 기준', ['설정·검증 시간 50% 단축 가능성 (GT-1)', '쾌적 유지 95% + 피크 전력 20%↓ (GT-2)', '추가 HW 없이 가전 2종 이상 협조 제어 (GT-3)', '유료 실증 또는 구매 조건을 제시한 파트너 1곳 이상'], C.okSoft, C.ok],
    ['6명이 하지 않는 것', ['배터리 셀·인버터 전체 개발(인증 파트너 HW 사용)', '국가별 영업·설치 운영(사업부·현지 파트너)', '미국 PCS·V2H·중동 B2G는 2028년 파트너와 확장'], C.panel, C.muted],
    ['특허 출원 후보 (2027)', ['직교 시험 신호로 다기기 배선·상 동시 식별 (GT-1)', '쾌적 이탈 확률 제약 flex 용량 계산 (GT-2)', '가전 제어기가 코디네이터가 되는 게이트웨이 없는 HEMS (GT-3)', '무작위 휴지 구간 기반 실시간 인과 절감 정산 (GT-4)', '열부하·EV 여유 공통 단위 환산 배분 (GT-5)'], C.copperSoft, C.copper]];
  cards.forEach((c, i) => {
    const x = 0.6 + (i % 2) * 6.13, y = 1.55 + Math.floor(i / 2) * 2.72;
    R(s, x, y, 5.93, 2.55, c[2]);
    T(s, c[0], x + 0.3, y + 0.18, 5.4, 0.35, { fontSize: 14.5, bold: true, color: c[3] });
    T(s, bullets(c[1]), x + 0.3, y + 0.62, 5.4, 1.9, { fontSize: 13, paraSpaceAfter: 4 });
  });
}

// ================= 평가표 =================
const s4 = section('s4');
{
  const t = tables(s4)[0];
  const head = t[0].map(c => plain(c.html)); const rows = t.slice(1);
  const pillCell = c => { const k = (c.html.match(/p-(hi|mid|lo)/) || [])[1] || 'mid'; return { text: plain(c.html), options: { bold: true, align: 'center', color: { hi: C.ok, mid: C.warn, lo: C.bad }[k], fill: { color: { hi: C.okSoft, mid: C.warnSoft, lo: C.badSoft }[k] } } }; };
  [[0, 6], [6, 12], [12, 17]].forEach(([a, b], k) => {
    const s = slide('4 · 신기술 후보 17개 평가', `평가표 (${k + 1}/3)`, 'HEMS = SW·AI 기술 · EP = Energy Pack HW 기술 · 굵은 GT 표시는 Gateway Technology 구성 기술', '');
    const body = [head.map(HDR)];
    rows.slice(a, b).forEach((r, i) => {
      const f0 = r[0].html; const name = plain(f0.replace(/<span[\s\S]*$/, ''));
      const layer = /lay h/.test(f0) ? 'HEMS' : 'EP'; const gt = (f0.match(/pill p-hi">(GT-\d)/) || [])[1];
      const fill = { fill: { color: i % 2 ? C.panel : C.white } };
      body.push([{ text: [{ text: name, options: { bold: true, breakLine: true } }, { text: layer + (gt ? '  ·  ' + gt : ''), options: { bold: true, color: layer === 'HEMS' ? C.blue : C.copper, fontSize: 11 } }], options: fill },
        ...r.slice(1, 8).map(c => ({ text: plain(c.html), options: fill })), pillCell(r[8]), pillCell(r[9])]);
    });
    s.addTable(body, { x: 0.6, y: 1.8, w: 12.13, colW: [1.95, 1.2, 0.95, 2.05, 1.05, 1.1, 1.2, 1.15, 0.7, 0.78], fontFace: F, fontSize: 12.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'top', margin: [4, 5, 4, 5], autoPage: false });
  });
}

// ================= 상세 =================
divider('PART 1-B', '기술별 상세', '17개 신기술 × 10개 질문 · Pain Point에서 특허 아이디어까지');
const s5 = section('s5');
const Q = ['1. 출발 Pain Point', '2. 기존 HEMS로 어려운 이유', '3. 핵심 기술', '4. AI가 필요한 이유', '5. HW 필요성', '6. 고객 Value', '7. Utility/Grid Value', '8. 경쟁사', '9. 격차 포인트', '10. 특허 아이디어'];
const evalRows = tables(s4)[0].slice(1);
const dets = [...s5.matchAll(/<details class="tech"[^>]*><summary><span class="id">(T\d+)<\/span><h3>([\s\S]*?)<\/h3>([\s\S]*?)<\/summary>[\s\S]*?<dl class="ten">([\s\S]*?)<\/dl>/g)];
dets.sort((a, b) => +a[1].slice(1) - +b[1].slice(1));
dets.forEach(d => {
  const id = d[1], name = plain(d[2]); const tag = plain(d[3]);
  const dd = [...d[4].matchAll(/<dd>([\s\S]*?)<\/dd>/g)].map(m => plain(m[1]));
  const ev = evalRows.find(r => plain(r[0].html).startsWith(id + ' '));
  const layer = ev && /lay h/.test(ev[0].html) ? 'HEMS' : 'EP';
  const s = slide('5 · 기술별 상세', id + '  ' + name, '', notesFor(d[0]));
  const accent = layer === 'HEMS' ? C.blue : C.copper;
  R(s, 0.6, 1.5, 2.9, 5.35, layer === 'HEMS' ? C.blueSoft : C.copperSoft);
  T(s, id, 0.85, 1.68, 2.4, 0.75, { fontSize: 38, bold: true, color: accent });
  let px = 0.85; px += pill(s, layer, px, 2.5, layer === 'HEMS' ? 'blue' : 'gt') + 0.12; if (tag) pill(s, tag, px, 2.5, tag.startsWith('GT') ? 'gt' : 'mid');
  if (ev) {
    const info = [['해결 Pain Point', plain(ev[1].html)], ['Target', plain(ev[2].html)], ['기존과 차이', plain(ev[3].html)], ['차별화 · 필요성', plain(ev[8].html) + ' · ' + plain(ev[9].html)]];
    let y = 3.05; info.forEach(([k, v]) => { T(s, k, 0.85, y, 2.45, 0.26, { fontSize: 9.5, bold: true, color: C.muted }); T(s, v, 0.85, y + 0.27, 2.45, 0.62, { fontSize: 11 }); y += 0.93; });
  }
  const rows = Q.map((q, i) => [{ text: q, options: { bold: true, color: C.muted, fill: { color: i % 2 ? C.white : C.panel } } }, { text: dd[i] || '', options: { fill: { color: i % 2 ? C.white : C.panel } } }]);
  s.addTable(rows, { x: 3.7, y: 1.5, w: 9.03, colW: [2.0, 7.03], rowH: 0.535, fontFace: F, fontSize: 11, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [3, 7, 3, 7], autoPage: false });
});

// ================= Part 2 =================
divider('PART 2', '왜 필요한가', '시장 장벽 · 국가별 Pain Point · 가설 검증 · 경쟁사 · 가치사슬');
const s6 = section('s6');
{
  const s = slide('6 · 왜 필요한가', '가장 큰 장벽은 Integration Tax(통합 비용)', '보조금·Net Metering이 사라진 뒤 가치는 여러 기기를 함께 움직일 때만 생기는데, 그 결합 비용이 가치보다 크다', notesFor(s6));
  const st = [['58%', '미국 주택 PV+ESS 비용 중 soft cost', 'Solar Builder †'], ['2–8k€', '독일 HP+PV HEMS 통합 비용', 'heatload †'], ['9,000명', '영국 MCS 인증 HP 설치자(2030년 연 45만 대 목표)', 'MCS Foundation'],
    ['49% → 17%', '영국: TOU 절약 기대 → 실제 전환 의향. 이유는 “절감 보장 없음”', 'Ofgem 2025'], ['12.9%', '여름 DR 이벤트 중 사용자가 제어를 끈 비율', 'ScienceDirect 2025'], ['108 / 500MW', '텍사스 가정 VPP 한도 대비 실제 적격 용량', 'ERCOT 2026.3']];
  st.forEach((a, i) => {
    const x = 0.6 + (i % 3) * 4.1, y = 1.95 + Math.floor(i / 3) * 2.45;
    R(s, x, y, 3.9, 2.25, C.panel);
    T(s, a[0], x + 0.3, y + 0.22, 3.4, 0.85, { fontSize: 34, bold: true, color: i < 3 ? C.copper : C.blue });
    T(s, a[1], x + 0.3, y + 1.1, 3.4, 0.7, { fontSize: 12.5 });
    T(s, a[2], x + 0.3, y + 1.85, 3.4, 0.25, { fontSize: 9.5, color: C.muted });
  });
}
{
  const bars = [...s6.matchAll(/<div class="bar"><span>(.*?)<\/span>.*?<span class="v">(\d+)\/12<\/span>/g)].map(m => [plain(m[1]), +m[2]]).reverse();
  const s = slide('6 · 왜 필요한가', '기술은 있는데 시장이 안 크는 이유', '12개 시장 Pain Point를 원인별로 분류 — 기술 “성능”이 아니라 기술이 현장에 닿는 경로에서 막힘', notesFor(s6));
  s.addChart(pres.charts.BAR, [{ name: '시장 수', labels: bars.map(b => b[0]), values: bars.map(b => b[1]) }], {
    x: 0.6, y: 1.85, w: 7.3, h: 4.95, barDir: 'bar', chartColors: [C.blue], showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 11, dataLabelColor: C.ink, dataLabelFontFace: F,
    catAxisLabelFontFace: F, catAxisLabelFontSize: 11, catAxisLabelColor: C.ink, valAxisHidden: true, valAxisMaxVal: 12, valAxisMinVal: 0,
    valGridLine: { style: 'none' }, catGridLine: { style: 'none' }, showLegend: false, showTitle: true, title: '근본 장벽이 나타난 시장 수 (12개 중)', titleFontFace: F, titleFontSize: 12, titleColor: C.muted });
  R(s, 8.2, 1.85, 4.53, 4.95, C.panel);
  T(s, '해석', 8.5, 2.05, 4, 0.35, { fontSize: 15, bold: true, color: C.copper });
  T(s, bullets(['예측·ESS·TOU 최적화, EV·HVAC 스케줄링은 이미 상용 기능', 'ROI 불확실과 설치·연동 복잡성이 상위 → 모두 “통합 비용” 문제', 'HEMS를 쓰게 만드는 것 자체(참여·신뢰)가 3번째 장벽', '계통 연계 용량·제어권 분절은 Energy Pack과 VPP의 장벽', '정전·설치 규제는 특정 지역 장벽']), 8.5, 2.55, 4, 4.1, { fontSize: 12.5, paraSpaceAfter: 9 });
}
{
  const [barrier, trig] = tables(s6);
  const s = slide('6 · 왜 필요한가', '장벽별로 기존 HEMS가 못 푸는 이유와 해결 기술', '', notesFor(s6));
  tableSlideFromRows(s, barrier, [2.4, 7.93, 1.8], { fs: 15 });
  const s_ = slide('6 · 왜 필요한가', '시장을 여는 Trigger와 성숙도', '시장 성숙도에 따라 들어갈 제품 형태가 다르다', notesFor(s6));
  tableSlideFromRows(s_, trig, [2.2, 2.0, 3.9, 4.03], { fs: 16 });
}

// ================= 국가 =================
const s7 = section('s7');
{
  const t = tables(s7)[0]; const head = t[0]; const rows = t.slice(1);
  const byName = k => rows.filter(r => k.some(x => plain(r[0].html).startsWith(x)));
  const groups = [['유럽 ①', '독일·영국: 규제는 앞서가는데 설치·연동과 사용 참여가 따라가지 못함', ['독일', '영국']],
    ['유럽 ②', '프랑스·네덜란드·남유럽: 요금·보상·기후가 바뀌어 기존 기기 운전이 어긋남', ['프랑스', '네덜란드', '스페인']],
    ['북미', '미국·California·Texas: 보조금 종료 후 패널 용량·수익 불확실·정전', ['미국', 'California', 'Texas']],
    ['중동', 'UAE·사우디·카타르: 소비자 ROI보다 냉방 피크와 Comfort', ['UAE', '사우디', '카타르']]];
  groups.forEach(g => {
    const rr = byName(g[2]); const seg = rr.map(r => r.map(c => c.html).join('')).join('');
    const s = slide('7 · 국가별 Pain Point · ' + g[0], g[1], '', notesFor(seg));
    const body = [head.map(c => HDR(plain(c.html)))];
    rr.forEach((r, i) => body.push(r.map((c, j) => ({ text: plain(c.html.replace(/<span class="k">/, '\n<span>')), options: { bold: j === 0 || j === 5, color: j === 5 ? C.blue : C.ink, fill: { color: i % 2 ? C.panel : C.white } } }))));
    s.addTable(body, { x: 0.6, y: 1.5, w: 12.13, colW: [1.05, 3.45, 3.2, 1.6, 1.78, 1.05], fontFace: F, fontSize: 13, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'top', margin: [3, 5, 3, 5], autoPage: false });
  });
}

// ================= 가설 =================
const s8 = section('s8');
{
  const t = tables(s8)[0];
  const s = slide('8 · 가설 검증', '검토 중이던 가설과 후보의 판정', '', notesFor(s8));
  const body = t.map((r, i) => r.map((c, j) => {
    if (i === 0) return HDR(plain(c.html));
    if (j === 3) { const k = (c.html.match(/p-(hi|mid|lo)/) || [])[1] || 'mid'; return { text: plain(c.html), options: { bold: true, align: 'center', color: { hi: C.ok, mid: C.warn, lo: C.bad }[k], fill: { color: { hi: C.okSoft, mid: C.warnSoft, lo: C.badSoft }[k] } } }; }
    return { text: plain(c.html), options: { bold: j === 0, fill: { color: i % 2 ? C.white : C.panel } } };
  }));
  s.addTable(body, { x: 0.6, y: 1.5, w: 12.13, colW: [2.3, 3.2, 3.63, 1.15, 1.85], fontFace: F, fontSize: 12.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [3, 6, 3, 6], autoPage: false });
}

// ================= 경쟁사 =================
const s9 = section('s9');
{
  const [mat, typ] = tables(s9);
  const s = slide('9 · 경쟁사', '누가 어디까지 하고 있나', '● 상용 제품   ◐ 일부 국가·파트너·제한적   · 없음/미확인', notesFor(s9));
  const body = [mat[0].map(c => HDR(plain(c.html)))];
  mat.slice(1).forEach((r, i) => body.push(r.map((c, j) => {
    const fill = { color: i % 2 ? C.panel : C.white };
    const d = (c.attr.match(/dot (f|h|n)/) || [])[1];
    if (d) return { text: plain(c.html), options: { align: 'center', fontSize: 11, color: d === 'f' ? C.blue : d === 'h' ? C.copper : 'B8C0C8', fill } };
    return { text: plain(c.html), options: { bold: j === 0, fill } };
  })));
  s.addTable(body, { x: 0.6, y: 1.75, w: 12.13, colW: [1.85, 0.52, 0.52, 0.52, 0.52, 0.58, 0.6, 0.52, 0.7, 5.8], rowH: 0.215, fontFace: F, fontSize: 9, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [1, 5, 1, 5], autoPage: false });
  const s_ = slide('9 · 경쟁사', 'HEMS 관점의 경쟁 구도와 우리가 노릴 자리', '', notesFor(s9));
  const b2 = typ.map((r, i) => r.map((c, j) => i === 0 ? HDR(plain(c.html)) : { text: plain(c.html), options: { bold: j === 0 || i === typ.length - 1, color: i === typ.length - 1 ? C.copper : C.ink, fill: { color: i === typ.length - 1 ? C.copperSoft : i % 2 ? C.white : C.panel } } }));
  s_.addTable(b2, { x: 0.6, y: 1.5, w: 12.13, colW: [1.9, 2.9, 2.2, 2.6, 2.53], fontFace: F, fontSize: 12, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [5, 6, 5, 6], autoPage: false });
  const gaps = [['공백 1', '냉난방 회사에는 전력·계통 스택이 없다', 'HP를 팔지만 유연성 수익화는 외부 HEMS에 맡긴다.'], ['공백 2', '전력기기 회사는 열부하·가전을 모른다', 'ESS 중심이라 냉난방·가전을 쾌적 보장과 함께 제어하지 못한다.'], ['공백 3', '게이트웨이 없는 중립 HEMS가 없다', '모두 자사 허브·요금·채널이 전제. 가전 자체가 HEMS가 되는 제품은 없다.']];
  gaps.forEach((g, i) => { const x = 0.6 + i * 4.1; R(s_, x, 5.3, 3.9, 1.6, C.panel); pill(s_, g[0], x + 0.25, 5.42, 'blue'); T(s_, g[1], x + 0.25, 5.8, 3.45, 0.4, { fontSize: 13, bold: true }); T(s_, g[2], x + 0.25, 6.2, 3.45, 0.65, { fontSize: 10.5, color: C.muted }); });
}

// ================= 가치사슬 =================
const s10 = section('s10');
{
  const rows = [...s10.matchAll(/<div class="chain( head)?">([\s\S]*?)<\/div>\s*(?=<div class="chain|<\/div>)/g)];
  const parsed = rows.map(r => [...(r[2] + '</div>').matchAll(/<div>([\s\S]*?)<\/div>/g)].map(m => m[1]));
  const s = slide('10 · 가치사슬', 'Country Pain Point → Business Opportunity', '기술에서 출발하지 않고 Pain Point에서 거꾸로 추적 — “이 기술이 없어서 지금 시장이 못 크고 있다”', notesFor(s10));
  const body = parsed.map((r, i) => r.map((c, j) => {
    if (i === 0) return HDR(plain(c));
    const h = (c.match(/<span class="h">(.*?)<\/span>/) || [])[1];
    const txt = plain(c.replace(/<span class="h">.*?<\/span>/, ''));
    const hw = i === 5;
    if (j === 0) return { text: [{ text: plain(h || ''), options: { bold: true, color: C.copper, fontSize: 9.5, breakLine: true } }, { text: txt }], options: { fill: { color: C.white } } };
    return { text: txt, options: { bold: j === 3, color: j === 3 ? (hw ? C.copper : C.blue) : C.ink, fill: { color: j === 3 ? (hw ? C.copperSoft : C.blueSoft) : (i % 2 ? C.panel : C.white) } } };
  }));
  s.addTable(body, { x: 0.6, y: 1.8, w: 12.13, colW: [2.05, 1.65, 1.8, 1.75, 1.65, 1.5, 1.73], fontFace: F, fontSize: 12.5, color: C.ink, border: { type: 'solid', pt: 0.5, color: C.line }, valign: 'middle', margin: [4, 5, 4, 5], autoPage: false });
}

// ================= 사업 구조 =================
const s11 = section('s11');
{
  const s = slide('11 · 사업 구조와 Region', '신기술이 들어가는 사업', '', notesFor(s11));
  const biz = [['사업 ①', 'Smart Cottage', '주거 공간 + 설비 결합 상품. Energy Pack과 Smart Home이 함께 들어가고 HEMS가 묶음. 미국 판매 제외.'], ['사업 ②', 'Energy Pack', 'PV·ESS·HP·AC(+EV 충전) 묶음. Smart Cottage에 포함되며 별도 사업으로도 전개.'], ['공통 기술', 'HEMS', 'Energy Pack·Smart Home·자사 가전 모두의 두뇌(GT-1~4). HW·가전 경쟁력과 서비스 수익으로 돈을 번다.']];
  biz.forEach((b, i) => { const x = 0.6 + i * 4.1; R(s, x, 1.5, 3.9, 1.72, i === 2 ? C.blueSoft : C.panel); T(s, b[0], x + 0.25, 1.62, 3.4, 0.25, { fontSize: 10, color: C.muted, bold: true }); T(s, b[1], x + 0.25, 1.88, 3.4, 0.4, { fontSize: 16, bold: true, color: i === 2 ? C.blue : C.ink }); T(s, b[2], x + 0.25, 2.33, 3.45, 0.85, { fontSize: 10.5 }); });
  tableSlideFromRows(s, tables(s11)[0], [1.4, 1.95, 2.05, 2.9, 2.63, 1.2], { fs: 12, y: 3.4 });
}

// ================= 회의 반영 =================
const s12 = section('s12');
{
  const s = slide('12 · 회의 반영', '회의에서 요구한 것과 대응', '', notesFor(s12));
  tableSlideFromRows(s, tables(s12)[0], [6.3, 5.83], { fs: 14.5, y: 1.5 });
  const s_ = slide('12 · 회의 반영', '이 전략이 틀릴 수 있는 조건과 일정', '', '');
  const cards = [['이 전략이 틀릴 수 있는 조건', ['Matter·EEBus 표준화가 빨라져 커미셔닝 난이도가 크게 낮아지는 경우(GT-1 가치 감소, GT-3는 유리)', '1KOMMA5·Octopus가 제조사 중립 레트로핏을 대규모 확산', '자사 가전 제어기의 연산·통신 여유가 내장 HEMS에 부족한 경우', '자사 HP·AC 내부 API 접근이 확보되지 않는 경우'], C.badSoft, C.bad],
    ['일정', ['아웃라인 리뷰 → 소장 리뷰(매주)', '10월 말 CTO 보고: 과제 지속·범위·인력 결정', '12월까지 2027–2028 로드맵 확정'], C.blueSoft, C.blue]];
  cards.forEach((c, i) => { const x = 0.6 + i * 6.13; R(s_, x, 1.55, 5.93, 4.6, c[2]); T(s_, c[0], x + 0.3, 1.75, 5.4, 0.4, { fontSize: 16, bold: true, color: c[3] }); T(s_, bullets(c[1]), x + 0.3, 2.3, 5.4, 3.7, { fontSize: 13.5, paraSpaceAfter: 10 }); });
}

// ================= 결정 요청 / Closing =================
{
  const s = slide('', '', '', '', true);
  T(s, 'CTO 결정 요청', 0.8, 0.9, 11, 0.4, { fontSize: 14, bold: true, color: C.copper, charSpacing: 2 });
  T(s, '예측 정확도가 아니라,\n연결되고 · 조절하고 · 가전 안에서 돌고 · 증명하는 HEMS', 0.8, 1.4, 11.7, 1.5, { fontSize: 30, bold: true, color: C.white });
  const asks = ['과제 지속, 신기술 범위를 HEMS GT-1~4 중심으로 재정의 (GT-5는 2027 하반기 설계부터)', 'PL 포함 6명 재배치', '자사 HP·AC·가전 내부 제어 API와 제어기 펌웨어 접근 승인', '12주 실증 뒤 계속/축소 판단 기준 합의: 설정 시간 50%↓ · 쾌적 95% + 피크 20%↓ · 가전 2종 협조 제어 · 유료 파트너 1곳'];
  asks.forEach((a, i) => {
    const y = 3.35 + i * 0.78;
    s.addShape(pres.shapes.OVAL, { x: 0.8, y, w: 0.5, h: 0.5, fill: { color: C.copper }, line: { color: C.copper } });
    T(s, String(i + 1), 0.8, y, 0.5, 0.5, { fontSize: 15, bold: true, color: C.white, align: 'center', valign: 'middle' });
    T(s, a, 1.55, y, 11, 0.5, { fontSize: 16, color: C.white, valign: 'middle' });
  });
}

// ================= 출처 =================
{
  const ids = Object.keys(SOURCES).map(Number).sort((a, b) => a - b);
  const dom = t => { const u = (t.match(/https?:\/\/([^/\s]+)/) || [])[1] || ''; return t.replace(/\s*—?\s*https?:\/\/\S+/g, '').replace(/\s*·\s*$/, '').trim() + (u ? ' · ' + u.replace(/^www\./, '') : ''); };
  const per = 28; const pages = Math.ceil(ids.length / per);
  for (let p = 0; p < pages; p++) {
    const chunk = ids.slice(p * per, (p + 1) * per);
    const s = slide('13 · 출처', `참고 자료 (${p + 1}/${pages})`, '2025–2026 자료 우선 · † 업계·2차 자료 · 전체 URL은 발표자 노트', chunk.map(i => `[${i}] ${SOURCES[i]}`).join('\n'));
    const half = Math.ceil(chunk.length / 2);
    [chunk.slice(0, half), chunk.slice(half)].forEach((col, j) => T(s, col.map(i => `${i}. ${dom(SOURCES[i])}`).join('\n'), 0.6 + j * 6.13, 1.8, 5.93, 5.1, { fontSize: 11, lineSpacingMultiple: 1.25, fit: 'shrink' }));
  }
}

pres.writeFile({ fileName: OUT }).then(f => console.log('written', f, 'slides', n));
