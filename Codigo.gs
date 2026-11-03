/**
 * Carteira do Coutinho — servidor (Google Apps Script)
 * Liga a folha "Carteira do Coutinho · Dados" à app e recebe os pagamentos Apple Pay.
 *
 * 1) Cola tudo isto em script.google.com → Novo projeto → Guardar.
 * 2) Escolhe a função "instalar" e carrega em Executar (autoriza quando pedir).
 * 3) Implementar → Nova implementação → Aplicação Web → Executar como: Eu → Acesso: Qualquer pessoa.
 * 4) Copia o link que acaba em /exec: vai para a app e para o Atalho do iPhone.
 * Atualizar (v3): cola este código por cima, Guardar, Executar "instalar", depois
 *    Implementar → Gerir implementações → ✎ → Versão: Nova versão → Implementar (o link /exec fica igual).
 */
const SHEET_ID = '1cGpGYWRqlLg2gOjhnH91M1pXNh8g9Csz_nmKeJIuE1c';
const TZ = 'Europe/Lisbon';
const COLS = {
  Movimentos: ['id','data','hora','tipo','valor','loja','categoria','conta','para','metodo','origem','rever','criado','cartao','morada','notas'],
  Contas: ['id','nome','tipo','saldo_inicial','data_inicial','criado_em','cor','ordem','oculta'],
  Categorias: ['id','nome','emoji','cor','tipo','limite','ordem'],
  Recorrentes: ['id','nome','tipo','valor','variavel','flex','dia','conta','para','categoria','metodo','proximo','ativo']
};
const BET = /betano|gobet|placard|betclic|bwin|solverde|esc ?online|casino|lebull|nossa ?aposta|betway|22bet|pokerstars|luckia|bacana|moosh|trustly/i;
const RULES = [
  [BET, 'apostas'],
  [/continente|pingo doce|lidl|auchan(?! energy)|mercadona|minipreco|intermarche|aldi|spar|froiz/, 'supermercado'],
  [/mcdonald|burger|kfc|telepizza|pizza|starbucks|cafe|pastelaria|padaria|restaurante|tasca|sushi|glovo|uber eats|bolt food|five guys|francesinha/, 'restaurantes'],
  [/galp|repsol|prio|cepsa|auchan energy|via verde|uber|bolt|metro|stcp|andante|comboios|parque|portagem|ratp|navigo|sncf/, 'transportes'],
  [/farmacia|wells|clinica|hospital|dentista|cuf|lusiadas|optica/, 'saude'],
  [/netflix|spotify|youtube|apple com|icloud|patreon|disney plus|hbo|prime video|chatgpt|anthropic|claude/, 'subscricoes'],
  [/padel|ginasio|fitness|decathlon|sport zone|whoop/, 'desporto'],
  [/cinema|nos lusomundo|uci|ticketline|blueticket|steam|playstation|nintendo|fc porto/, 'lazer'],
  [/fnac|worten|zara|primark|amazon|ikea|shein|bershka|mango|leroy|aliexpress|temu|pcdiga|h m/, 'compras'],
  [/booking|airbnb|hotel|ryanair|easyjet|tap air|euro disney/, 'viagens'],
  [/edp|endesa|meo|altice|vodafone|nos comunicacoes|aguas/, 'casa']
];

function book() { return SpreadsheetApp.openById(SHEET_ID); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function cell(v, k) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, k === 'hora' ? 'HH:mm' : 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function read(name) {
  const v = book().getSheetByName(name).getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = cell(r[i], k)); return o; });
}
function parse(s) { try { return JSON.parse(s || '{}'); } catch (e) { return {}; } }
function jogo() { const v = book().getSheetByName('Jogo').getRange('A2:C2').getValues()[0]; return { estado: parse(v[0]), config: parse(v[1]), precos: parse(v[2]) }; }
function hoje() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim(); }

function doGet(e) {
  const a = e && e.parameter ? e.parameter.a : '';
  if (a === 'all') {
    const j = jogo();
    return json({ ok: true, tx: read('Movimentos'), contas: read('Contas'), cats: read('Categorias'), recur: read('Recorrentes'), jogo: j.estado, config: j.config, precos: j.precos, agora: new Date().toISOString() });
  }
  return ContentService.createTextOutput('Carteira ativa ✔');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const raw = e && e.postData ? e.postData.contents : '';
    const d = parse(raw);
    if (!d.a) return json(applePay(d, raw));             // vem do Atalho do iPhone
    if (d.a === 'put') { upsert('Movimentos', d.row); return json({ ok: true }); }
    if (d.a === 'del') { remove('Movimentos', d.id); return json({ ok: true }); }
    if (d.a === 'jogo') { book().getSheetByName('Jogo').getRange('A2').setValue(JSON.stringify(d.jogo)); return json({ ok: true }); }
    if (d.a === 'config') { book().getSheetByName('Jogo').getRange('B2').setValue(JSON.stringify(d.config)); return json({ ok: true }); }
    if (d.a === 'conta') { upsert('Contas', d.row); return json({ ok: true }); }
    if (d.a === 'cat') { upsert('Categorias', d.row); return json({ ok: true }); }
    if (d.a === 'recur') { upsert('Recorrentes', d.row); return json({ ok: true }); }
    if (d.a === 'recurdel') { remove('Recorrentes', d.id); return json({ ok: true }); }
    if (d.a === 'quick') return json(rapido(String(d.texto || '')));   // vem do Atalho "Carteira" da Siri
    return json({ ok: false, erro: 'ação desconhecida' });
  } finally { lock.releaseLock(); }
}

function findRow(sh, id) { const f = sh.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext(); return f ? f.getRow() : 0; }
function upsert(name, row) {
  const sh = book().getSheetByName(name), cols = COLS[name];
  const vals = cols.map(k => row[k] === undefined || row[k] === null ? '' : row[k]);
  const r = findRow(sh, row.id);
  if (r) sh.getRange(r, 1, 1, cols.length).setValues([vals]); else sh.appendRow(vals);
}
function remove(name, id) { const sh = book().getSheetByName(name); const r = findRow(sh, id); if (r && r > 1) sh.deleteRow(r); }

function valor(v) {
  let s = String(v || '').replace(/[^\d,.]/g, ''); if (!s) return 0;
  const dec = Math.max(s.lastIndexOf(','), s.lastIndexOf('.'));
  let n = (dec >= 0 && s.length - dec - 1 <= 2 && s.length - dec - 1 > 0) ? parseFloat(s.slice(0, dec).replace(/[.,]/g, '') + '.' + s.slice(dec + 1)) : parseFloat(s.replace(/[.,]/g, ''));
  return isFinite(n) ? Math.round(n * 100) / 100 : 0;
}
function categoria(loja) {
  const n = norm(loja);
  if (BET.test(n)) return 'apostas';
  const hist = read('Movimentos').filter(t => t.tipo === 'out' && t.rever !== '1' && norm(t.loja) === n);
  if (hist.length) return hist[hist.length - 1].categoria;
  for (const [re, c] of RULES) if (re.test(' ' + n + ' ')) return c;
  return 'outros';
}
function applePay(d, raw) {
  if (!String(d.loja || '').trim() && !valor(d.valor)) return { ok: true, ignorado: 'sem dados da transação' }; // atalho corrido sem pagamento
  const now = new Date(), loja = String(d.loja || '').trim() || 'Apple Pay', v = valor(d.valor);
  const map = (jogo().config.cardMap) || {}, card = norm(d.cartao || d.nome || '');
  const conta = map[card] || (/revolut/.test(card) ? 'rev' : /bpi/.test(card) ? 'bpi' : /trade/.test(card) ? 'tr' : /212/.test(card) ? 't212' : /klarna/.test(card) ? 'klarna' : 'rev');
  const id = 'ap_' + Utilities.getUuid().replace(/-/g, '').slice(0, 14);
  upsert('Movimentos', { id: id, data: hoje(), hora: Utilities.formatDate(now, TZ, 'HH:mm'), tipo: 'out', valor: v, loja: loja,
    categoria: categoria(loja), conta: conta, para: '', metodo: 'Apple Pay', origem: 'applepay', rever: '1', criado: String(now.getTime()),
    cartao: String(d.cartao || d.nome || ''), morada: [d.morada, d.cidade].filter(Boolean).join(', '), notas: v ? '' : 'Valor não reconhecido: ' + String(raw).slice(0, 80) });
  return { ok: true, id: id };
}

/* Siri: "Ei Siri, Carteira" → dita "5 euros café BPI" → fica registado */
function eur(c) { return (c < 0 ? '−' : '') + Math.abs(c).toFixed(2).replace('.', ',') + ' €'; }
function rapido(texto) {
  const src = String(texto).replace(/(\d)\s*euros?\s*e\s*(\d{1,2})(\s*c[eê]ntimos?)?/i, '$1,$2').replace(/€/g, ' ');
  const mm = src.match(/\d+(?:[.,]\d{1,2})?/), v = mm ? valor(mm[0]) : 0;
  if (!v) return { ok: false, msg: 'Não percebi o valor. Diz por exemplo: 5 euros café BPI.' };
  const isIn = /^\s*\+|receb|ganhei|entrou|entraram|salario|ordenado/.test(norm(src));
  // palavras originais (para manter acentos) + versão normalizada (para comparar)
  let tk = src.replace(mm[0], ' ').split(/\s+/).filter(Boolean).map(w => ({ w: w.replace(/^[^\wÀ-ÿ]+|[^\wÀ-ÿ]+$/g, ''), n: norm(w) })).filter(x => x.n);
  const take = phrase => { const p = phrase.split(' '); for (let i = 0; i + p.length <= tk.length; i++) if (p.every((x, j) => tk[i + j].n === x)) { tk.splice(i, p.length); return true; } return false; };
  const contas = read('Contas').filter(c => c.oculta !== '1').map(c => ({ id: c.id, nome: c.nome, n: norm(c.nome) })).sort((a, b) => b.n.length - a.n.length);
  let conta = '';
  for (const c of contas) if (take(c.n)) { conta = c.id; break; }
  if (!conta) for (const c of contas) { const w = c.n.split(' ')[0]; if (w.length > 2 && take(w)) { conta = c.id; break; } }
  const METODOS = [['mb way', 'MB WAY'], ['mbway', 'MB WAY'], ['apple pay', 'Apple Pay'], ['multibanco', 'Cartão'], ['cartao', 'Cartão'], ['dinheiro', 'Dinheiro'], ['online', 'Online'], ['transferencia', 'Transferência']];
  let metodo = '';
  for (const [k, m] of METODOS) if (take(k)) { metodo = m; break; }
  if (!metodo && conta === 'cash') metodo = 'Dinheiro';
  const FILLER = new Set('gastei paguei comprei recebi ganhei entrou entraram euro euros eur centimo centimos no na nos nas num numa em com pela pelo a o os as de do da dos das um uma e para pra hoje conta cartao carteira'.split(' '));
  tk = tk.filter(x => !FILLER.has(x.n));
  const tipo = isIn ? 'in' : 'out';
  let loja = tk.map(x => x.w).join(' ').trim();
  loja = loja ? loja.charAt(0).toUpperCase() + loja.slice(1) : (isIn ? 'Entrada' : 'Compra');
  const mov = read('Movimentos'), n = norm(loja);
  const hist = mov.filter(x => x.tipo === tipo && norm(x.loja) === n).pop(), last = mov.filter(x => x.tipo === tipo && x.origem !== 'extrato').pop();
  const rec = read('Recorrentes').filter(r => r.tipo === tipo).find(r => n && norm(r.nome).split(' ').some(w => w.length > 2 && n.split(' ').indexOf(w) >= 0));
  const cats = read('Categorias'), catOk = id => cats.some(c => c.id === id);
  let cat = hist ? hist.categoria : rec ? rec.categoria : tipo === 'out' ? categoria(loja) : 'outras-entradas';
  if (!catOk(cat)) cat = tipo === 'out' ? 'outros' : (catOk('outras-entradas') ? 'outras-entradas' : (cats.find(c => c.tipo === 'in') || {}).id || '');
  conta = conta || (hist && hist.conta) || (rec && rec.conta) || (last && last.conta) || 'rev';
  metodo = metodo || (hist && hist.metodo) || (tipo === 'in' ? 'Transferência' : 'Cartão');
  const now = new Date(), id = 's_' + Utilities.getUuid().replace(/-/g, '').slice(0, 14);
  upsert('Movimentos', { id: id, data: hoje(), hora: Utilities.formatDate(now, TZ, 'HH:mm'), tipo: tipo, valor: v, loja: loja, categoria: cat, conta: conta, para: '',
    metodo: metodo, origem: 'siri', rever: '', criado: String(now.getTime()), cartao: '', morada: '', notas: '' });
  const cn = (cats.find(c => c.id === cat) || {}).nome || cat, an = (contas.find(c => c.id === conta) || {}).nome || conta;
  return { ok: true, id: id, msg: (tipo === 'out' ? 'Gasto registado: ' : 'Entrada registada: ') + eur(v) + ', ' + loja + ', ' + cn + ', ' + an + '.' };
}

/* Recorrentes fixos: lançados todos os dias às 6h como "moedas por apanhar" */
function proxima(data, dia) {
  const p = data.split('-').map(Number); let y = p[0], m = p[1] + 1; if (m > 12) { m = 1; y++; }
  const max = new Date(y, m, 0).getDate(); const d = Math.min(dia || p[2], max);
  return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
}
function diario() {
  const sh = book().getSheetByName('Movimentos'), h = hoje();
  read('Recorrentes').forEach(r => {
    if (r.ativo !== '1' || r.variavel === '1' || r.flex === '1' || !r.proximo) return;
    let next = r.proximo, k = 0;
    while (next <= h && k < 24) {
      const id = 'r_' + r.id + '_' + next;
      if (!findRow(sh, id)) upsert('Movimentos', { id: id, data: next, hora: '', tipo: r.tipo, valor: Number(r.valor), loja: r.nome, categoria: r.categoria,
        conta: r.conta, para: r.para, metodo: r.metodo, origem: 'recorrente', rever: '1', criado: String(Date.now()), cartao: '', morada: '', notas: '' });
      next = proxima(next, Number(r.dia)); k++;
    }
    if (next !== r.proximo) { r.proximo = next; upsert('Recorrentes', r); }
  });
}
/* Investimentos: preço das ETFs atualizado de hora a hora (Yahoo Finance; se falhar, GOOGLEFINANCE) */
function yahoo(sym) {
  const r = UrlFetchApp.fetch('https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(sym) + '?range=1mo&interval=1d',
    { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (r.getResponseCode() !== 200) return null;
  const res = (JSON.parse(r.getContentText()).chart.result || [])[0]; if (!res) return null;
  const out = { d: {}, p: Number(res.meta.regularMarketPrice) || null, cur: res.meta.currency };
  const ts = res.timestamp || [], cl = ((res.indicators.quote || [])[0] || {}).close || [];
  ts.forEach((t, i) => { if (cl[i] != null) out.d[Utilities.formatDate(new Date(t * 1000), TZ, 'yyyy-MM-dd')] = Math.round(cl[i] * 10000) / 10000; });
  return out;
}
function gfinance(g) {
  const c = book().getSheetByName('Jogo').getRange('D2');
  c.setFormula('=GOOGLEFINANCE("' + g + '")'); SpreadsheetApp.flush(); Utilities.sleep(1500);
  const v = Number(c.getValue()); c.clearContent(); return isFinite(v) && v > 0 ? v : null;
}
function precos() {
  const j = jogo(), inv = (j.config.invest || []), sh = book().getSheetByName('Jogo'), h = j.precos || {}, d = hoje();
  const legs = {}; inv.forEach(i => (i.legs || []).forEach(l => legs[l.y] = l.g));
  Object.keys(legs).forEach(y => {
    let r = null; try { r = yahoo(y); } catch (e) { r = null; }
    const cur = h[y] || { d: {} };
    if (r && r.p && (!r.cur || r.cur === 'EUR')) { Object.assign(cur.d, r.d); cur.p = r.p; cur.at = new Date().toISOString(); cur.src = 'yahoo'; }
    else { const p = legs[y] ? gfinance(legs[y]) : null; if (p) { cur.d[d] = p; cur.p = p; cur.at = new Date().toISOString(); cur.src = 'google'; } }
    if (cur.p && !cur.d[d] && Utilities.formatDate(new Date(), TZ, 'u') < 6) cur.d[d] = cur.p;
    const ks = Object.keys(cur.d).sort(); ks.slice(0, Math.max(0, ks.length - 400)).forEach(k => delete cur.d[k]);
    h[y] = cur;
  });
  sh.getRange('C1').setValue('precos'); sh.getRange('C2').setValue(JSON.stringify(h));
}
function instalar() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('diario').timeBased().everyDays(1).atHour(6).inTimezone(TZ).create();
  ScriptApp.newTrigger('precos').timeBased().everyHours(1).create();
  diario(); precos();
  Logger.log('Pronto! A Carteira está instalada. Agora faz Implementar → Nova implementação → Aplicação Web.');
}
