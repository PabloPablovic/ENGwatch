/* Registratúra – Outlook doplnok (testovacia verzia)
 * Číta vybraný e-mail v kontexte prihláseného používateľa a pošle údaje do Power Automate (HTTP trigger).
 * V kóde NIE SÚ žiadne adresy ani kľúče – používateľ ich zadá v Nastaveniach (Office roamingSettings).
 */
'use strict';

const ZNACKY_TTL_MS = 24 * 60 * 60 * 1000;      // obnova číselníka raz denne
const LS_ZNACKY = 'reg.znacky';                  // cache číselníka (localStorage tohto doplnku)
const LS_SENDERS = 'reg.senders';                // posledná Adresa/Značka pre e-mailovú adresu
const LS_CTYPE = 'reg.ctype';                    // ktorý Content-Type prešiel cez CORS

const $ = (id) => document.getElementById(id);
let znacky = [];          // [{id, t}]
let ctx = {};             // údaje z aktuálneho e-mailu

/* ---------- pomocné ---------- */
function showMsg(text, kind) {
  const m = $('msg');
  m.textContent = text;
  m.className = 'msg ' + (kind || '');
  m.hidden = !text;
}
function lsGet(key, fallback) {
  try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; }
}
function lsSet(key, val) {
  try { localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* bez cache to tiež funguje */ }
}
function fmtAddr(a) {
  if (!a) return '';
  return a.displayName && a.displayName !== a.emailAddress ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress;
}
function getSettings() {
  const rs = Office.context.roamingSettings;
  return { url: rs.get('flowUrl') || '', key: rs.get('flowKey') || '' };
}

/* ---------- volanie flowu ---------- */
async function callFlow(payload) {
  const s = getSettings();
  if (!s.url || !s.key) throw new Error('Chýbajú nastavenia (adresa flowu a kľúč).');
  const body = JSON.stringify(Object.assign({ key: s.key }, payload));
  // application/json spôsobí CORS preflight; ak ho flow nepodporí, skúsime text/plain (bez preflightu)
  const order = lsGet(LS_CTYPE, null) === 'text/plain' ? ['text/plain', 'application/json'] : ['application/json', 'text/plain'];
  let lastErr;
  for (const ct of order) {
    try {
      const r = await fetch(s.url, { method: 'POST', headers: { 'Content-Type': ct }, body });
      const text = await r.text();
      let data;
      try { data = JSON.parse(text); } catch (e) { data = { ok: false, error: text || ('HTTP ' + r.status) }; }
      if (!r.ok || data.ok === false) throw Object.assign(new Error(data.error || ('HTTP ' + r.status)), { http: true });
      lsSet(LS_CTYPE, ct);
      return data;
    } catch (e) {
      lastErr = e;
      if (e.http) break;           // server odpovedal chybou – nemá zmysel skúšať iný Content-Type
    }
  }
  throw lastErr;
}

/* ---------- číselník Značiek ---------- */
async function loadZnacky(force) {
  const cache = lsGet(LS_ZNACKY, null);
  if (!force && cache && Date.now() - cache.ts < ZNACKY_TTL_MS && cache.items.length) {
    znacky = cache.items;
    $('znackaInfo').textContent = `${znacky.length} značiek (z cache)`;
  } else {
    $('znackaInfo').textContent = 'Načítavam číselník…';
    const d = await callFlow({ action: 'getZnacky' });
    znacky = (d.znacky || []).map((z) => ({ id: Number(z.id), t: String(z.t) }))
      .sort((a, b) => a.t.localeCompare(b.t, 'sk'));
    lsSet(LS_ZNACKY, { ts: Date.now(), items: znacky });
    $('znackaInfo').textContent = `${znacky.length} značiek`;
  }
  renderZnacky();
}
function renderZnacky(selectId) {
  const sel = $('fZnacka');
  const keep = selectId != null ? selectId : Number(sel.value) || null;
  const q = $('fZnackaFilter').value.trim().toLowerCase();
  const words = q ? q.split(/\s+/) : [];
  sel.textContent = '';
  const none = new Option('— bez značky —', '');
  sel.add(none);
  for (const z of znacky) {
    const t = z.t.toLowerCase();
    if (words.every((w) => t.includes(w))) sel.add(new Option(z.t, String(z.id)));
  }
  if (keep && [...sel.options].some((o) => o.value === String(keep))) sel.value = String(keep);
  else sel.value = '';
}

/* ---------- načítanie e-mailu ---------- */
function loadItem() {
  const item = Office.context.mailbox.item;
  $('done').hidden = true;
  showMsg('');
  if (!item) { showMsg('Vyber e-mail.', 'err'); $('form').hidden = true; return; }

  const me = (Office.context.mailbox.userProfile.emailAddress || '').toLowerCase();
  const from = item.from || item.sender;
  const sent = from && from.emailAddress && from.emailAddress.toLowerCase() === me;
  const counterpart = sent ? (item.to || []).map(fmtAddr).join('; ') : fmtAddr(from);
  const firstAddr = sent ? ((item.to || [])[0] || {}).emailAddress : (from || {}).emailAddress;

  ctx = {
    messageId: item.internetMessageId || '',
    datumEmailu: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : '',
    senderKey: (firstAddr || '').toLowerCase(),
  };

  $('fSmer').value = sent ? 'odoslaný' : 'doručený';
  $('fEmail').value = counterpart;
  $('fPredmet').value = item.subject || '';
  $('fPridelene').value = Office.context.mailbox.userProfile.emailAddress || '';
  $('fPoznamka').value = '';
  $('fZnackaFilter').value = '';

  // návrh z predošlého evidovania rovnakej adresy (len lokálne v tomto PC)
  const hist = lsGet(LS_SENDERS, {})[ctx.senderKey] || {};
  $('fAdresa').value = hist.adresa || '';
  renderZnacky(hist.znackaId || null);

  const emlOk = Office.context.requirements.isSetSupported('Mailbox', '1.14');
  $('emlWrap').hidden = !emlOk;
  $('fEml').checked = emlOk;

  $('form').hidden = false;
  $('btnSubmit').disabled = false;
}

function getEml() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item.getAsFileAsync) return resolve(null);
    item.getAsFileAsync((res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : null));
  });
}

/* ---------- odoslanie ---------- */
async function submit(ev) {
  ev.preventDefault();
  const btn = $('btnSubmit');
  btn.disabled = true;
  showMsg('Odosielam…');
  try {
    const znackaId = Number($('fZnacka').value) || null;
    const data = {
      smer: $('fSmer').value,
      emailovaAdresa: $('fEmail').value.trim(),
      adresa: $('fAdresa').value.trim(),
      predmet: $('fPredmet').value.trim(),
      znackaId,
      pridelene: $('fPridelene').value.trim(),
      poznamka: $('fPoznamka').value.trim(),
      datumEmailu: ctx.datumEmailu,
      messageId: ctx.messageId,
      zaevidoval: Office.context.mailbox.userProfile.emailAddress || '',
    };
    if (!data.emailovaAdresa) throw new Error('Emailová adresa je povinná.');

    let eml = null;
    if (!$('emlWrap').hidden && $('fEml').checked) {
      const b64 = await getEml();
      if (b64) {
        const name = (data.predmet || 'email').replace(/[\\/:*?"<>|\r\n\t]+/g, ' ').trim().slice(0, 80) || 'email';
        eml = { name: name + '.eml', base64: b64 };
      }
    }

    const r = await callFlow({ action: 'create', data, eml });

    const senders = lsGet(LS_SENDERS, {});
    senders[ctx.senderKey] = { adresa: data.adresa, znackaId };
    lsSet(LS_SENDERS, senders);

    showMsg('');
    $('form').hidden = true;
    $('done').hidden = false;
    $('doneCislo').textContent = r.cislo || ('ID ' + r.id);
    const link = $('doneLink');
    if (r.url) { link.href = r.url; link.parentElement.hidden = false; } else { link.parentElement.hidden = true; }
  } catch (e) {
    showMsg('Chyba: ' + e.message, 'err');
    btn.disabled = false;
  }
}

/* ---------- nastavenia ---------- */
function openSettings() {
  const s = getSettings();
  $('setUrl').value = s.url;
  $('setKey').value = s.key;
  $('settings').hidden = false;
  $('form').hidden = true;
  $('done').hidden = true;
}
function saveSettings() {
  const url = $('setUrl').value.trim();
  const key = $('setKey').value.trim();
  if (!/^https:\/\//i.test(url)) { showMsg('Adresa musí začínať https://', 'err'); return; }
  const rs = Office.context.roamingSettings;
  rs.set('flowUrl', url);
  rs.set('flowKey', key);
  rs.saveAsync((res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) { showMsg('Nastavenia sa nepodarilo uložiť.', 'err'); return; }
    $('settings').hidden = true;
    showMsg('Nastavenia uložené.', 'ok');
    init(true);
  });
}

/* ---------- štart ---------- */
async function init(forceZnacky) {
  const s = getSettings();
  if (!s.url || !s.key) { openSettings(); return; }
  loadItem();
  try { await loadZnacky(forceZnacky); } catch (e) {
    $('znackaInfo').textContent = '';
    showMsg('Číselník sa nenačítal: ' + e.message, 'err');
  }
}

Office.onReady(() => {
  $('form').addEventListener('submit', submit);
  $('fZnackaFilter').addEventListener('input', () => renderZnacky());
  $('btnReloadZnacky').addEventListener('click', () => loadZnacky(true).catch((e) => showMsg(e.message, 'err')));
  $('btnSettings').addEventListener('click', openSettings);
  $('btnSaveSettings').addEventListener('click', saveSettings);
  $('btnCancelSettings').addEventListener('click', () => { $('settings').hidden = true; init(false); });
  // pripnutý panel: pri zmene vybraného e-mailu znova načítaj
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { if ($('settings').hidden) loadItem(); });
  }
  init(false);
});
