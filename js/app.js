/* Mælaborð Þjórsárskóla
 * Les Excel-gagnaskjal (einn flipi á hvern lið) og teiknar mælaborð.
 * Engin gögn eru send neitt: skjalið er lesið í vafranum.
 */
(function () {
  'use strict';

  const CFG = window.MAELABORD_CONFIG || {};

  /* ------------------------------------------------------------------ */
  /* Snið                                                                */
  /* ------------------------------------------------------------------ */
  const nf = (o) => new Intl.NumberFormat('is-IS', o);
  const F = {
    int: (v) => (v == null ? '—' : nf({ maximumFractionDigits: 0 }).format(v)),
    dec: (v) => (v == null ? '—' : nf({ maximumFractionDigits: 1 }).format(v)),
    pct: (v) => (v == null ? '—' : nf({ style: 'percent', maximumFractionDigits: 1 }).format(v)),
    pct0: (v) => (v == null ? '—' : nf({ style: 'percent', maximumFractionDigits: 0 }).format(v)),
    mkr: (v) => (v == null ? '—' : nf({ minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(v)),
    sign: (v, f) => (v == null ? '—' : (v > 0 ? '+' : v < 0 ? '−' : '±') + f(Math.abs(v))),
  };
  const signInt = (v) => F.sign(v, F.int);
  const signPct = (v) => F.sign(v, (x) => nf({ style: 'percent', minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(x));
  const signPP = (v) => (v == null ? '—' : F.sign(v * 100, (x) => nf({ maximumFractionDigits: 1 }).format(x)) + ' pr.st.');

  const MANUDIR = [8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6];
  const MAN = ['', 'jan', 'feb', 'mar', 'apr', 'maí', 'jún', 'júl', 'ágú', 'sep', 'okt', 'nóv', 'des'];
  const MANL = ['', 'janúar', 'febrúar', 'mars', 'apríl', 'maí', 'júní', 'júlí', 'ágúst', 'september', 'október', 'nóvember', 'desember'];

  /* ------------------------------------------------------------------ */
  /* Lestur Excel                                                        */
  /* ------------------------------------------------------------------ */
  const norm = (s) => String(s == null ? '' : s).trim().toLowerCase().replace(/\s+/g, ' ').replace(/[–—]/g, '-');
  let vantar = [];

  function lesaFlipa(wb, nafn) {
    const lykill = wb.SheetNames.find((n) => norm(n) === norm(nafn));
    if (!lykill) { vantar.push(nafn); return []; }
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[lykill], { header: 1, raw: true, defval: null, blankrows: false });
    if (!rows.length) return [];
    const hdr = rows[0].map(norm);
    return rows.slice(1)
      .map((r) => { const o = {}; hdr.forEach((h, i) => { if (h) o[h] = r[i]; }); return o; })
      .filter((o) => Object.values(o).some((v) => v !== null && v !== ''));
  }
  const c = (o, nafn) => o[norm(nafn)];

  function tala(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/\s/g, '');
    const prosent = s.includes('%');
    const n = Number(s.replace('%', '').replace(/\.(?=\d{3}(\D|$))/g, '').replace(',', '.'));
    if (!isFinite(n)) return null;
    return prosent ? n / 100 : n;
  }
  const hlutfall = (v) => { const n = tala(v); return n == null ? null : n > 1.5 ? n / 100 : n; };

  function manudur(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 864e5));
      return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1 };
    }
    if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1 };
    const s = String(v).trim();
    let r = s.match(/^(\d{4})[-./](\d{1,2})/);
    if (r) return { y: +r[1], m: +r[2] };
    r = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})/);
    if (r) return { y: +r[3], m: +r[2] };
    r = s.match(/^(\d{1,2})[./-](\d{4})$/);
    if (r) return { y: +r[2], m: +r[1] };
    const idx = MAN.findIndex((x) => x && s.toLowerCase().startsWith(x));
    r = s.match(/(\d{4})/);
    if (idx > 0 && r) return { y: +r[1], m: idx };
    return null;
  }
  const skolaarUr = (ym) => (ym.m >= 8 ? `${ym.y}-${ym.y + 1}` : `${ym.y - 1}-${ym.y}`);
  function skolaar(v) {
    if (v == null || v === '') return null;
    const r = String(v).match(/(\d{4})/);
    if (!r) return null;
    const y = +r[1];
    return `${y}-${y + 1}`;
  }
  const fyrraAr = (sy) => { const y = +sy.slice(0, 4); return `${y - 1}-${y}`; };
  const arMerki = (sy) => sy.slice(0, 4) + '–' + sy.slice(7, 9);
  const yfirstandandi = (sy) => new Date() < new Date(+sy.slice(0, 4) + 1, 6, 1);

  function manadarflipi(wb, nafn, dalkar, fn) {
    return lesaFlipa(wb, nafn).map((o) => {
      const ym = manudur(c(o, 'Mánuður'));
      if (!ym) return null;
      const r = { y: ym.y, m: ym.m, sy: skolaarUr(ym) };
      dalkar.forEach(([k, h]) => { r[k] = (fn || tala)(c(o, h)); });
      return r;
    }).filter(Boolean);
  }

  function lesaGogn(buf) {
    vantar = [];
    const wb = XLSX.read(buf, { type: 'array' });
    const D = {};
    D.stillingar = {};
    lesaFlipa(wb, 'Stillingar').forEach((o) => { D.stillingar[norm(c(o, 'Stilling'))] = c(o, 'Gildi'); });
    D.nemendur = lesaFlipa(wb, 'Nemendur').map((o) => ({
      sy: skolaar(c(o, 'Skólaár')), s1: tala(c(o, '1.–4. bekkur')), s2: tala(c(o, '5.–7. bekkur')), s3: tala(c(o, '8.–10. bekkur')),
      isat: tala(c(o, 'ÍSAT-nemendur')), stod: tala(c(o, 'Stuðningur stoðþjónustu')),
    })).filter((r) => r.sy);
    D.maeting = manadarflipi(wb, 'Mæting', [['v', 'Mæting']], hlutfall);
    D.forfoll = manadarflipi(wb, 'Forföll', [['dagar', 'Forfalladagar'], ['mannad', 'Mannaðir dagar']]);
    D.starfsfolk = lesaFlipa(wb, 'Starfsfólk').map((o) => ({
      sy: skolaar(c(o, 'Skólaár')), kennarar: tala(c(o, 'Kennarar')), rett: tala(c(o, 'Með kennsluréttindi')),
      art: tala(c(o, 'ART-þjálfun lokið')), bofs: tala(c(o, 'BOFS-þjálfun lokið')),
    })).filter((r) => r.sy);
    D.samskipti = manadarflipi(wb, 'Samskipti', [['post', 'Tölvupóstar til foreldra'], ['heim', 'Fréttir á heimasíðu'], ['fjol', 'Fréttir til fjölmiðla']]);
    D.oryggi = manadarflipi(wb, 'Öryggi', [['vedur', 'Röskun vegna veðurs'], ['aflyst', 'Skóla aflýst'], ['atvik', 'Atvik og slys'], ['aefing', 'Æfingar']]);
    D.akstur = manadarflipi(wb, 'Skólaakstur', [['fellt', 'Ferðir felldar niður'], ['atvik', 'Atvik'], ['kvort', 'Kvartanir']]);
    D.stundir = lesaFlipa(wb, 'Kennslustundir').map((o) => ({
      sy: skolaar(c(o, 'Skólaár')), arg: tala(c(o, 'Árgangur')), ut: tala(c(o, 'Útinám')),
      list: tala(c(o, 'List- og verkgreinar')), ithr: tala(c(o, 'Íþróttir')),
    })).filter((r) => r.sy && r.arg);
    D.innramat = lesaFlipa(wb, 'Innra mat').map((o) => ({
      sy: skolaar(c(o, 'Skólaár')), svid: c(o, 'Svið'), metid: tala(c(o, 'Metið')) || 0,
      vinnsla: tala(c(o, 'Í vinnslu')) || 0, ekki: tala(c(o, 'Ekki hafið')) || 0,
    })).filter((r) => r.sy && r.svid);
    D.fjarhagur = lesaFlipa(wb, 'Fjárhagur').map((o) => ({
      ar: tala(c(o, 'Ár')), timabil: c(o, 'Tímabil'), flokkur: c(o, 'Flokkur'),
      aars: tala(c(o, 'Áætlun ársins')), atil: tala(c(o, 'Áætlun til dags')), rtil: tala(c(o, 'Raun til dags')),
    })).filter((r) => r.ar && r.flokkur);
    D.skraningar = lesaFlipa(wb, 'Skráningar').map((o) => {
      const ym = manudur(c(o, 'Dagsetning'));
      return ym ? { sy: skolaarUr(ym), tegund: norm(c(o, 'Tegund')) } : null;
    }).filter((r) => r && r.tegund);
    D.arsm = lesaFlipa(wb, 'Ársmælingar').map((o) => ({
      sy: skolaar(c(o, 'Skólaár')), heiti: norm(c(o, 'Mælikvarði')), gildi: c(o, 'Gildi'),
    })).filter((r) => r.sy && r.heiti);
    D.maelikvardar = lesaFlipa(wb, 'Mælikvarðar').map((o) => ({
      kodi: String(c(o, 'Kóði') || '').trim(), heiti: c(o, 'Mælikvarði'), svid: c(o, 'Svið') || 'Annað', uppruni: c(o, 'Uppruni') || '',
      eining: norm(c(o, 'Eining')) || 'fjoldi', vidmid: tala(c(o, 'Viðmið')), att: norm(c(o, 'Átt')) === 'niður' ? 'nidur' : 'upp',
      tegund: norm(c(o, 'Tegund')) === 'ná' ? 'na' : 'halda', lind: norm(c(o, 'Gagnalind')), virkur: norm(c(o, 'Virkur')) !== 'nei',
    })).filter((r) => r.heiti);
    D.maelikvardar.forEach((m) => { if (m.eining === 'pct' && m.vidmid != null && m.vidmid > 1.5) m.vidmid /= 100; });
    D.vstund = lesaFlipa(wb, 'Viðmiðunarstundaskrá').map((o) => ({
      stig: c(o, 'Stig'), fra: tala(c(o, 'Frá árgangi')), til: tala(c(o, 'Til árgangs')),
      list: tala(c(o, 'List- og verkgreinar')), ithr: tala(c(o, 'Íþróttir')), ut: tala(c(o, 'Útinám')),
    })).filter((r) => r.fra && r.til);
    return D;
  }

  /* ------------------------------------------------------------------ */
  /* Útreikningar                                                        */
  /* ------------------------------------------------------------------ */
  let D = null;
  const st = { ar: null, tab: 'yfirlit', svid: 'Allt' };

  const iAri = (arr, sy) => arr.filter((r) => r.sy === sy);
  const summa = (arr, k) => arr.reduce((s, r) => s + (r[k] || 0), 0);
  const medal = (arr, k) => { const v = arr.map((r) => r[k]).filter((x) => x != null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  const sidasti = (arr, sy) => iAri(arr, sy).sort((a, b) => (a.y * 12 + a.m) - (b.y * 12 + b.m)).pop() || null;
  function manadarkort(arr, sy) {
    const m = new Map();
    iAri(arr, sy).forEach((r) => m.set(r.m, r));
    return m;
  }
  function samaManudi(arr, sy) {
    const nu = iAri(arr, sy);
    const man = new Set(nu.map((r) => r.m));
    return { nu, adur: iAri(arr, fyrraAr(sy)).filter((r) => man.has(r.m)) };
  }
  const stigFyrir = (arg) => D.vstund.find((s) => arg >= s.fra && arg <= s.til);
  const vidmidArgangs = (arg, k) => {
    const s = stigFyrir(arg); if (!s || s[k] == null) return null;
    return k === 'ut' ? s.ut : s[k] / (s.til - s.fra + 1);
  };
  function nemendurAlls(sy) {
    const r = D.nemendur.find((x) => x.sy === sy);
    return r ? (r.s1 || 0) + (r.s2 || 0) + (r.s3 || 0) : null;
  }
  function fjarhagsar(sy) {
    const y = +sy.slice(0, 4);
    if (D.fjarhagur.some((r) => r.ar === y)) return y;
    const ar = D.fjarhagur.map((r) => r.ar).filter((a) => a <= y + 1).sort();
    return ar.length ? ar[ar.length - 1] : null;
  }

  const REIKNAD = {
    maeting: (sy) => medal(iAri(D.maeting, sy), 'v'),
    forfoll_mannad: (sy) => { const r = iAri(D.forfoll, sy); const d = summa(r, 'dagar'); return d ? summa(r, 'mannad') / d : null; },
    akstur_atvik: (sy) => { const r = iAri(D.akstur, sy); return r.length ? (summa(r, 'atvik') + summa(r, 'kvort')) / r.length : null; },
    frettir_heimasida: (sy) => medal(iAri(D.samskipti, sy), 'heim'),
    aefingar: (sy) => { const r = iAri(D.oryggi, sy); return r.length ? summa(r, 'aefing') : null; },
    utinam_argangar: (sy) => {
      const r = iAri(D.stundir, sy); if (!r.length) return null;
      return r.filter((x) => x.ut != null && x.ut >= (vidmidArgangs(x.arg, 'ut') || 0)).length / r.length;
    },
    list_verk_stig: (sy) => stigHlutfall(sy, 'list'),
    ithrottir_stig: (sy) => stigHlutfall(sy, 'ithr'),
    art: (sy) => kennaraHlutfall(sy, 'art'),
    bofs: (sy) => kennaraHlutfall(sy, 'bofs'),
    leyfisbref: (sy) => kennaraHlutfall(sy, 'rett'),
    rekstur_fravik: (sy) => {
      const y = +sy.slice(0, 4); const r = D.fjarhagur.filter((x) => x.ar === y);
      const a = summa(r, 'atil'); return a ? summa(r, 'rtil') / a - 1 : null;
    },
    innramat_metid: (sy) => {
      const r = iAri(D.innramat, sy); const alls = summa(r, 'metid') + summa(r, 'vinnsla') + summa(r, 'ekki');
      return alls ? summa(r, 'metid') / alls : null;
    },
  };
  function stigHlutfall(sy, k) {
    const r = iAri(D.stundir, sy); if (!r.length || !D.vstund.length) return null;
    const nad = D.vstund.filter((s) => summa(r.filter((x) => x.arg >= s.fra && x.arg <= s.til), k) >= (s[k] || 0)).length;
    return nad / D.vstund.length;
  }
  function kennaraHlutfall(sy, k) {
    const r = D.starfsfolk.find((x) => x.sy === sy);
    return r && r.kennarar ? (r[k] || 0) / r.kennarar : null;
  }

  function gildiMaelikvarda(m, sy) {
    if (m.lind === 'reiknað') return REIKNAD[m.kodi] ? REIKNAD[m.kodi](sy) : null;
    const nafn = norm(m.heiti), kodi = norm(m.kodi);
    if (m.lind === 'skráningar') {
      const r = iAri(D.skraningar, sy);
      if (!r.length) return null;
      return r.filter((x) => x.tegund === nafn || x.tegund === kodi).length;
    }
    const r = D.arsm.filter((x) => x.sy === sy && (x.heiti === nafn || x.heiti === kodi)).pop();
    if (!r) return null;
    return m.eining === 'pct' ? hlutfall(r.gildi) : tala(r.gildi);
  }

  function skorkort(sy) {
    return D.maelikvardar.filter((m) => m.virkur).map((m) => {
      const nu = gildiMaelikvarda(m, sy), adur = gildiMaelikvarda(m, fyrraAr(sy));
      let stada = 'none';
      if (nu != null && m.vidmid != null) {
        const nad = m.att === 'nidur' ? nu <= m.vidmid + 1e-9 : nu >= m.vidmid - 1e-9;
        stada = nad ? 'ok' : m.tegund === 'na' && yfirstandandi(sy) ? 'prog' : 'bad';
      }
      return Object.assign({}, m, { nu, adur, stada });
    });
  }
  const STADA = { ok: ['Innan viðmiðs', 'st-ok'], prog: ['Í vinnslu', 'st-prog'], bad: ['Utan viðmiðs', 'st-bad'], none: ['Engin mæling', 'st-none'] };
  function snidMk(m, v) {
    if (v == null) return '—';
    if (m.eining === 'frav') return signPct(v);
    if (m.eining === 'pct') return F.pct(v);
    return F.dec(v);
  }
  const vidmidTexti = (m) => (m.vidmid == null ? '—' :
    (m.att === 'nidur' ? '≤ ' : '≥ ') + (m.eining === 'fjoldi' ? F.dec(m.vidmid) : F.pct0(m.vidmid)) + (m.tegund === 'na' ? ' (ná)' : ' (halda)'));
  const finnaMk = (kodi) => D.maelikvardar.find((m) => m.kodi === kodi && m.virkur);

  /* ------------------------------------------------------------------ */
  /* Teikning                                                            */
  /* ------------------------------------------------------------------ */
  const $ = (s) => document.querySelector(s);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const litur = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  const LITIR = () => ({ c1: litur('--c1'), c2: litur('--c2'), c3: litur('--c3'), c4: litur('--c4'), grey: litur('--grey'), bad: litur('--bad'), ref: litur('--ref'), grid: litur('--grid'), muted: litur('--muted') });
  const grofin = {};

  function kpi(container, lbl, val, sub, tgt, tone) {
    const k = el('div', 'kpi');
    k.appendChild(el('span', 'lbl', lbl));
    const v = el('span', 'val', val); if (tone) v.classList.add(tone); k.appendChild(v);
    if (sub) { const s = el('span', 'sub'); if (Array.isArray(sub)) sub.forEach((p) => s.appendChild(typeof p === 'string' ? document.createTextNode(p) : p)); else s.textContent = sub; k.appendChild(s); }
    if (tgt) k.appendChild(el('span', 'tgt', tgt));
    container.appendChild(k);
  }
  const breyting = (txt, gott) => { const s = el('span', gott == null ? '' : gott ? 'good' : 'bad', txt); return s; };
  const toneMk = (m, v) => {
    if (!m || v == null || m.vidmid == null) return null;
    return (m.att === 'nidur' ? v <= m.vidmid + 1e-9 : v >= m.vidmid - 1e-9) ? 'good' : 'bad';
  };

  function graf(id, config, hefurGogn, tomtTexti) {
    if (grofin[id]) { grofin[id].destroy(); delete grofin[id]; }
    const canvas = document.getElementById(id);
    const box = canvas.parentElement;
    box.querySelectorAll('.empty').forEach((n) => n.remove());
    if (!hefurGogn) {
      canvas.style.display = 'none';
      box.appendChild(el('div', 'empty', tomtTexti || 'Engin gögn skráð fyrir þetta skólaár.'));
      return;
    }
    canvas.style.display = '';
    const L = LITIR();
    const o = config.options = config.options || {};
    o.responsive = true; o.maintainAspectRatio = false;
    o.animation = { duration: 300 };
    o.plugins = Object.assign({ legend: { position: 'bottom', labels: { boxWidth: 12, boxHeight: 12 } } }, o.plugins || {});
    o.plugins.tooltip = Object.assign({ mode: 'index', intersect: false }, o.plugins.tooltip || {});
    o.scales = o.scales || {};
    ['x', 'y'].forEach((ax) => {
      o.scales[ax] = Object.assign({ grid: { color: L.grid }, border: { display: false } }, o.scales[ax] || {});
    });
    grofin[id] = new Chart(canvas, config);
  }
  const yTicks = (f) => ({ ticks: { callback: (v) => f(v) } });
  const tipLabel = (f) => ({ callbacks: { label: (ctx) => {
    const v = ctx.chart.options.indexAxis === 'y' ? ctx.parsed.x : ctx.parsed.y;
    return v == null ? null : ` ${ctx.dataset.label}: ${f(v)}`;
  } } });
  const vidmidsLina = (label, data, L, stepped) => ({
    type: 'line', label, data, borderColor: L.ref, borderWidth: 2, borderDash: [6, 4], pointRadius: 0,
    pointHoverRadius: 0, fill: false, stepped: stepped ? 'middle' : false, order: 0,
  });
  const manadarMerki = MANUDIR.map((m) => MAN[m]);
  const manadarRod = (arr, sy, k) => { const mk = manadarkort(arr, sy); return MANUDIR.map((m) => (mk.has(m) ? mk.get(m)[k] : null)); };
  const manadarTip = { callbacks: { title: (items) => MANL[MANUDIR[items[0].dataIndex]] } };
  const sidustuAr = (arr, sy, n) => Array.from(new Set(arr.map((r) => r.sy))).filter((y) => y <= sy).sort().slice(-n);

  /* ---------- Yfirlit ---------- */
  function teiknaYfirlit(sy) {
    const box = $('#kpis'); box.textContent = '';
    const nu = nemendurAlls(sy), adur = nemendurAlls(fyrraAr(sy));
    kpi(box, 'Nemendur', F.int(nu), nu != null && adur != null ? [breyting(signInt(nu - adur)), ' frá fyrra ári'] : null);

    const mM = finnaMk('maeting'); const ma = REIKNAD.maeting(sy); const sid = sidasti(D.maeting, sy);
    kpi(box, 'Mæting nemenda, meðaltal', F.pct(ma), sid ? `Síðast: ${MANL[sid.m]} ${F.pct(sid.v)}` : null,
      mM ? 'Viðmið ' + vidmidTexti(mM) : null, toneMk(mM, ma));

    const mF = finnaMk('forfoll_mannad'); const fo = REIKNAD.forfoll_mannad(sy);
    const sf = samaManudi(D.forfoll, sy); const fa = summa(sf.adur, 'dagar') ? summa(sf.adur, 'mannad') / summa(sf.adur, 'dagar') : null;
    kpi(box, 'Forföll sem tókst að manna', F.pct0(fo), fo != null && fa != null ? [breyting(signPP(fo - fa), fo >= fa), ' miðað við sömu mánuði í fyrra'] : null,
      mF ? 'Viðmið ' + vidmidTexti(mF) : null, toneMk(mF, fo));

    const im = iAri(D.innramat, sy); const ima = REIKNAD.innramat_metid(sy);
    kpi(box, 'Innra mat, gæðaviðmið metin', F.pct0(ima), im.length ? `${F.int(summa(im, 'metid'))} af ${F.int(summa(im, 'metid') + summa(im, 'vinnsla') + summa(im, 'ekki'))} viðmiðum` : null);

    const fy = fjarhagsar(sy); const fr = D.fjarhagur.filter((r) => r.ar === fy);
    const frav = summa(fr, 'atil') ? summa(fr, 'rtil') / summa(fr, 'atil') - 1 : null;
    kpi(box, `Rekstur miðað við áætlun${fy ? ' ' + fy : ''}`, signPct(frav), fr.length ? `${F.mkr(summa(fr, 'rtil'))} af ${F.mkr(summa(fr, 'atil'))} millj. kr.` : null,
      null, frav == null ? null : frav <= 0.005 ? 'good' : 'bad');

    const sa = samaManudi(D.akstur, sy);
    const ak = summa(sa.nu, 'atvik') + summa(sa.nu, 'kvort'), aka = summa(sa.adur, 'atvik') + summa(sa.adur, 'kvort');
    kpi(box, 'Atvik og kvartanir í skólaakstri', sa.nu.length ? F.int(ak) : '—',
      sa.nu.length && sa.adur.length ? [breyting(signInt(ak - aka), ak <= aka), ' miðað við sömu mánuði í fyrra'] : null);

    // skorkort
    const rows = skorkort(sy);
    const tal = { ok: 0, prog: 0, bad: 0, none: 0 }; rows.forEach((r) => tal[r.stada]++);
    const cnt = $('#skorkort-talning'); cnt.textContent = '';
    ['ok', 'prog', 'bad', 'none'].forEach((k) => {
      const b = el('div', 'count ' + STADA[k][1]); b.appendChild(el('b', null, F.int(tal[k]))); b.appendChild(el('span', null, STADA[k][0])); cnt.appendChild(b);
    });
    $('#yfirlit-texti').textContent =
      `Af ${F.int(rows.length)} mælikvörðum eru ${F.int(tal.ok)} innan viðmiðs, ${F.int(tal.prog)} í vinnslu og ${F.int(tal.bad)} utan viðmiðs. ` +
      (tal.none ? `${F.int(tal.none)} hafa ekki verið mældir enn á skólaárinu.` : '');
    $('#skorkort-skyring').textContent = yfirstandandi(sy)
      ? 'Skólaárið stendur yfir: markmið sem á að ná fyrir lok skólaárs („ná“) teljast í vinnslu þar til þau nást.'
      : 'Skólaárinu er lokið: markmið sem náðust ekki teljast utan viðmiðs.';
    $('#sk-fyrra').textContent = arMerki(fyrraAr(sy));
    $('#sk-nu').textContent = arMerki(sy) + (yfirstandandi(sy) ? ' til þessa' : '');

    const svid = ['Allt'].concat(Array.from(new Set(rows.map((r) => r.svid))));
    if (!svid.includes(st.svid)) st.svid = 'Allt';
    const chips = $('#svid-val'); chips.textContent = '';
    svid.forEach((s) => {
      const b = el('button', null, s); b.type = 'button'; b.setAttribute('aria-pressed', String(s === st.svid));
      b.addEventListener('click', () => { st.svid = s; teiknaYfirlit(st.ar); });
      chips.appendChild(b);
    });
    const tb = $('#skorkort tbody'); tb.textContent = '';
    let sidastaSvid = null;
    rows.filter((r) => st.svid === 'Allt' || r.svid === st.svid).forEach((r) => {
      if (st.svid === 'Allt' && r.svid !== sidastaSvid) {
        const g = el('tr', 'group'); const td = el('td', null, r.svid); td.colSpan = 7; g.appendChild(td); tb.appendChild(g);
        sidastaSvid = r.svid;
      }
      const tr = el('tr', 'st-row');
      tr.appendChild(el('td', null, r.heiti));
      tr.appendChild(el('td', 'svid hide-sm', r.svid));
      tr.appendChild(el('td', 'upp hide-sm', r.uppruni));
      tr.appendChild(el('td', 'num', snidMk(r, r.adur)));
      tr.appendChild(el('td', 'num', snidMk(r, r.nu)));
      tr.appendChild(el('td', 'num', vidmidTexti(r)));
      const td = el('td'); td.appendChild(el('span', 'pill ' + STADA[r.stada][1], STADA[r.stada][0])); tr.appendChild(td);
      tb.appendChild(tr);
    });
  }

  /* ---------- Nemendur ---------- */
  function teiknaNemendur(sy) {
    const L = LITIR();
    const box = $('#kpis-nemendur'); box.textContent = '';
    const r = D.nemendur.find((x) => x.sy === sy); const alls = nemendurAlls(sy); const adur = nemendurAlls(fyrraAr(sy));
    kpi(box, 'Nemendur alls', F.int(alls), alls != null && adur != null ? [breyting(signInt(alls - adur)), ' frá fyrra ári'] : null);
    kpi(box, 'ÍSAT-nemendur', r ? F.int(r.isat) : '—', r && alls ? F.pct0(r.isat / alls) + ' nemenda' : null);
    kpi(box, 'Njóta stuðnings stoðþjónustu', r ? F.int(r.stod) : '—', r && alls ? F.pct0(r.stod / alls) + ' nemenda' : null);

    const ar = sidustuAr(D.nemendur, sy, 6); const rr = ar.map((y) => D.nemendur.find((x) => x.sy === y));
    graf('c-nemendur', {
      type: 'bar',
      data: { labels: ar.map(arMerki), datasets: [
        { label: '1.–4. bekkur', data: rr.map((x) => x.s1), backgroundColor: L.c1, borderRadius: 3 },
        { label: '5.–7. bekkur', data: rr.map((x) => x.s2), backgroundColor: L.c2, borderRadius: 3 },
        { label: '8.–10. bekkur', data: rr.map((x) => x.s3), backgroundColor: L.c3, borderRadius: 3 },
      ] },
      options: { scales: { x: { stacked: true }, y: { stacked: true } }, plugins: { tooltip: Object.assign(tipLabel(F.int), { callbacks: Object.assign(tipLabel(F.int).callbacks, { footer: (it) => 'Alls: ' + F.int(it.reduce((s, i) => s + i.parsed.y, 0)) }) }) } },
    }, ar.length > 0);

    const as = ar.filter((y) => nemendurAlls(y));
    const pr = (k) => as.map((y) => { const x = D.nemendur.find((n) => n.sy === y); return x[k] == null ? null : x[k] / nemendurAlls(y); });
    graf('c-studningur', {
      type: 'line',
      data: { labels: as.map(arMerki), datasets: [
        { label: 'Stuðningur stoðþjónustu', data: pr('stod'), borderColor: L.c1, backgroundColor: L.c1, borderWidth: 2, pointRadius: 3, tension: 0 },
        { label: 'ÍSAT-nemendur', data: pr('isat'), borderColor: L.c2, backgroundColor: L.c2, borderWidth: 2, pointRadius: 3, tension: 0 },
      ] },
      options: { scales: { y: Object.assign({ beginAtZero: true }, yTicks(F.pct0)) }, plugins: { tooltip: { callbacks: { label: (ctx) => {
        const x = D.nemendur.find((n) => n.sy === as[ctx.dataIndex]); const k = ctx.datasetIndex === 0 ? 'stod' : 'isat';
        return ` ${ctx.dataset.label}: ${F.int(x[k])} (${F.pct0(ctx.parsed.y)})`;
      } } } } },
    }, as.length > 0);
  }

  /* ---------- Mæting ---------- */
  function teiknaMaeting(sy) {
    const L = LITIR(); const m = finnaMk('maeting');
    const box = $('#kpis-maeting'); box.textContent = '';
    const nu = iAri(D.maeting, sy); const ma = medal(nu, 'v'); const sid = sidasti(D.maeting, sy);
    const lag = nu.slice().sort((a, b) => a.v - b.v)[0];
    const sm = samaManudi(D.maeting, sy); const madur = medal(sm.adur, 'v');
    kpi(box, 'Meðalmæting skólaársins', F.pct(ma), ma != null && madur != null ? [breyting(signPP(ma - madur), ma >= madur), ' miðað við sömu mánuði í fyrra'] : null,
      m ? 'Viðmið ' + vidmidTexti(m) : null, toneMk(m, ma));
    kpi(box, 'Síðasti mánuður', sid ? F.pct(sid.v) : '—', sid ? MANL[sid.m] + ' ' + sid.y : null, null, toneMk(m, sid && sid.v));
    kpi(box, 'Lægsti mánuður', lag ? F.pct(lag.v) : '—', lag ? MANL[lag.m] + ' ' + lag.y : null, null, toneMk(m, lag && lag.v));

    const a = manadarRod(D.maeting, sy, 'v'), b = manadarRod(D.maeting, fyrraAr(sy), 'v');
    const allt = a.concat(b).filter((x) => x != null).concat(m && m.vidmid != null ? [m.vidmid] : []);
    const min = allt.length ? Math.max(0, Math.floor((Math.min(...allt) - 0.01) / 0.02) * 0.02) : 0;
    const ds = [
      { label: arMerki(sy), data: a, backgroundColor: L.c1, borderRadius: 3, order: 1 },
      { label: arMerki(fyrraAr(sy)), data: b, backgroundColor: L.grey, borderRadius: 3, order: 2 },
    ];
    if (m && m.vidmid != null) ds.push(vidmidsLina('Viðmið', MANUDIR.map(() => m.vidmid), L));
    graf('c-maeting', {
      type: 'bar', data: { labels: manadarMerki, datasets: ds },
      options: { scales: { y: Object.assign({ min, max: 1 }, { ticks: { stepSize: 0.02, callback: (v) => F.pct0(v) } }) }, plugins: { tooltip: Object.assign(tipLabel(F.pct), { callbacks: Object.assign(tipLabel(F.pct).callbacks, manadarTip.callbacks) }) } },
    }, a.some((x) => x != null) || b.some((x) => x != null));
  }

  /* ---------- Kennsla ---------- */
  function teiknaKennsla(sy) {
    const L = LITIR();
    const box = $('#kpis-kennsla'); box.textContent = '';
    const r = iAri(D.stundir, sy).sort((a, b) => a.arg - b.arg);
    const ut = r.filter((x) => x.ut >= (vidmidArgangs(x.arg, 'ut') || 0)).length;
    const n = D.vstund.length;
    kpi(box, 'Árgangar sem ná viðmiði í útinámi', r.length ? `${ut} af ${r.length}` : '—', null, null, r.length ? (ut === r.length ? 'good' : 'bad') : null);
    const lv = REIKNAD.list_verk_stig(sy), it = REIKNAD.ithrottir_stig(sy);
    kpi(box, 'Stig sem ná viðmiði í list- og verkgreinum', lv == null ? '—' : `${Math.round(lv * n)} af ${n}`, 'Viðmiðunarstundaskrá aðalnámskrár', null, lv == null ? null : lv >= 1 ? 'good' : 'bad');
    kpi(box, 'Stig sem ná viðmiði í íþróttum', it == null ? '—' : `${Math.round(it * n)} af ${n}`, 'Viðmiðunarstundaskrá aðalnámskrár', null, it == null ? null : it >= 1 ? 'good' : 'bad');
    const im = REIKNAD.innramat_metid(sy);
    kpi(box, 'Gæðaviðmið metin í innra mati', F.pct0(im));

    const argGraf = (id, k, lit, heiti) => {
      const vm = r.map((x) => vidmidArgangs(x.arg, k));
      graf(id, {
        type: 'bar',
        data: { labels: r.map((x) => x.arg + '.'), datasets: [
          { label: heiti, data: r.map((x) => x[k]), backgroundColor: r.map((x, i) => (vm[i] != null && x[k] < vm[i] ? L.bad : lit)), borderRadius: 3, order: 1 },
          vidmidsLina('Viðmið', vm.map((v) => (v == null ? null : Math.round(v))), L, true),
        ] },
        options: { plugins: { legend: { display: false }, tooltip: Object.assign(tipLabel((v) => F.int(v) + ' mín.'), {
          callbacks: Object.assign(tipLabel((v) => F.int(v) + ' mín.').callbacks, { title: (it) => r[it[0].dataIndex].arg + '. bekkur' }) }) },
          scales: { x: { ticks: { maxRotation: 0, autoSkip: false } }, y: { beginAtZero: true } } },
      }, r.length > 0, 'Stundaskrá ekki skráð fyrir þetta skólaár.');
    };
    argGraf('c-utinam', 'ut', L.c1, 'Útinám');
    argGraf('c-list', 'list', L.c2, 'List- og verkgreinar');
    argGraf('c-ithrottir', 'ithr', L.c3, 'Íþróttir');

    const ir = iAri(D.innramat, sy);
    graf('c-innramat', {
      type: 'bar',
      data: { labels: ir.map((x) => x.svid), datasets: [
        { label: 'Metið', data: ir.map((x) => x.metid), backgroundColor: L.c1, borderRadius: 3 },
        { label: 'Í vinnslu', data: ir.map((x) => x.vinnsla), backgroundColor: L.c3, borderRadius: 3 },
        { label: 'Ekki hafið', data: ir.map((x) => x.ekki), backgroundColor: L.grey, borderRadius: 3 },
      ] },
      options: { indexAxis: 'y', scales: { x: { stacked: true, beginAtZero: true }, y: { stacked: true, grid: { display: false } } }, plugins: { tooltip: tipLabel(F.int) } },
    }, ir.length > 0);
  }

  /* ---------- Starfsfólk ---------- */
  function teiknaStarfsfolk(sy) {
    const L = LITIR();
    const box = $('#kpis-starfsfolk'); box.textContent = '';
    const mF = finnaMk('forfoll_mannad'); const fo = REIKNAD.forfoll_mannad(sy);
    kpi(box, 'Forföll sem tókst að manna', F.pct0(fo), `${F.int(summa(iAri(D.forfoll, sy), 'dagar'))} forfalladagar á skólaárinu`, mF ? 'Viðmið ' + vidmidTexti(mF) : null, toneMk(mF, fo));
    [['leyfisbref', 'rett', 'Kennarar með kennsluréttindi'], ['art', 'art', 'Hafa lokið ART-þjálfun'], ['bofs', 'bofs', 'Hafa lokið BOFS-þjálfun']].forEach(([kodi, k, lbl]) => {
      const m = finnaMk(kodi); const v = kennaraHlutfall(sy, k); const r = D.starfsfolk.find((x) => x.sy === sy);
      kpi(box, lbl, F.pct0(v), r ? `${F.int(r[k])} af ${F.int(r.kennarar)} kennurum` : null, m ? 'Viðmið ' + vidmidTexti(m) : null, toneMk(m, v));
    });

    const mannad = manadarRod(D.forfoll, sy, 'mannad'), dagar = manadarRod(D.forfoll, sy, 'dagar');
    graf('c-forfoll', {
      type: 'bar',
      data: { labels: manadarMerki, datasets: [
        { label: 'Mannað', data: mannad, backgroundColor: L.c1, borderRadius: 3 },
        { label: 'Ekki mannað', data: dagar.map((d, i) => (d == null ? null : d - (mannad[i] || 0))), backgroundColor: L.bad, borderRadius: 3 },
      ] },
      options: { scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } }, plugins: { tooltip: Object.assign(tipLabel(F.int), { callbacks: Object.assign(tipLabel(F.int).callbacks, manadarTip.callbacks) }) } },
    }, dagar.some((x) => x != null));

    const ar = sidustuAr(D.starfsfolk, sy, 6); const rr = ar.map((y) => D.starfsfolk.find((x) => x.sy === y));
    const lina = (lbl, k, col) => ({ label: lbl, data: rr.map((x) => (x.kennarar ? x[k] / x.kennarar : null)), borderColor: col, backgroundColor: col, borderWidth: 2, pointRadius: 3, tension: 0 });
    graf('c-thjalfun', {
      type: 'line',
      data: { labels: ar.map(arMerki), datasets: [lina('Kennsluréttindi', 'rett', L.c1), lina('ART-þjálfun', 'art', L.c2), lina('BOFS-þjálfun', 'bofs', L.c3)] },
      options: { scales: { y: Object.assign({ min: 0, max: 1 }, yTicks(F.pct0)) }, plugins: { tooltip: tipLabel(F.pct0) } },
    }, ar.length > 0);
  }

  /* ---------- Samskipti ---------- */
  function teiknaSamskipti(sy) {
    const L = LITIR();
    const box = $('#kpis-samskipti'); box.textContent = '';
    const s = samaManudi(D.samskipti, sy);
    const mH = finnaMk('frettir_heimasida');
    [['post', 'Tölvupóstar til foreldra'], ['heim', 'Fréttir á heimasíðu'], ['fjol', 'Fréttir til fjölmiðla']].forEach(([k, lbl]) => {
      const nu = summa(s.nu, k), ad = summa(s.adur, k);
      kpi(box, lbl + ' á skólaárinu', s.nu.length ? F.int(nu) : '—', s.nu.length && s.adur.length ? [breyting(signInt(nu - ad), nu >= ad), ' miðað við sömu mánuði í fyrra'] : null,
        k === 'heim' && mH ? `Viðmið ${vidmidTexti(mH)} á mánuði · nú ${F.dec(medal(s.nu, 'heim'))}` : null);
    });
    const ds = [['post', 'Tölvupóstar til foreldra', L.c1], ['heim', 'Fréttir á heimasíðu', L.c2], ['fjol', 'Fréttir til fjölmiðla', L.c3]]
      .map(([k, lbl, col]) => ({ label: lbl, data: manadarRod(D.samskipti, sy, k), backgroundColor: col, borderRadius: 3 }));
    graf('c-samskipti', {
      type: 'bar', data: { labels: manadarMerki, datasets: ds },
      options: { scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { tooltip: Object.assign(tipLabel(F.int), { callbacks: Object.assign(tipLabel(F.int).callbacks, manadarTip.callbacks) }) } },
    }, s.nu.length > 0);
  }

  /* ---------- Öryggi ---------- */
  function teiknaOryggi(sy) {
    const L = LITIR();
    const box = $('#kpis-oryggi'); box.textContent = '';
    const r = iAri(D.oryggi, sy);
    const mA = finnaMk('aefingar');
    kpi(box, 'Dagar með röskun vegna veðurs', r.length ? F.int(summa(r, 'vedur')) : '—');
    kpi(box, 'Dagar sem skóla var aflýst', r.length ? F.int(summa(r, 'aflyst')) : '—');
    kpi(box, 'Skráð atvik og slys', r.length ? F.int(summa(r, 'atvik')) : '—');
    kpi(box, 'Rýmingar- og viðbragðsæfingar', r.length ? F.int(summa(r, 'aefing')) : '—', null, mA ? 'Viðmið ' + vidmidTexti(mA) + ' á ári' : null,
      r.length && mA && summa(r, 'aefing') >= mA.vidmid ? 'good' : null);
    const tt = Object.assign(tipLabel(F.int), { callbacks: Object.assign(tipLabel(F.int).callbacks, manadarTip.callbacks) });
    graf('c-vedur', {
      type: 'bar',
      data: { labels: manadarMerki, datasets: [
        { label: 'Röskun vegna veðurs', data: manadarRod(D.oryggi, sy, 'vedur'), backgroundColor: L.c1, borderRadius: 3 },
        { label: 'Skóla aflýst', data: manadarRod(D.oryggi, sy, 'aflyst'), backgroundColor: L.bad, borderRadius: 3 },
      ] },
      options: { scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } }, plugins: { tooltip: tt } },
    }, r.length > 0);
    graf('c-atvik', {
      type: 'bar',
      data: { labels: manadarMerki, datasets: [
        { label: 'Atvik og slys', data: manadarRod(D.oryggi, sy, 'atvik'), backgroundColor: L.c3, borderRadius: 3 },
        { label: 'Æfingar', data: manadarRod(D.oryggi, sy, 'aefing'), backgroundColor: L.c2, borderRadius: 3 },
      ] },
      options: { scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }, plugins: { tooltip: tt } },
    }, r.length > 0);
  }

  /* ---------- Skólaakstur ---------- */
  function teiknaAkstur(sy) {
    const L = LITIR();
    const box = $('#kpis-akstur'); box.textContent = '';
    const s = samaManudi(D.akstur, sy); const mA = finnaMk('akstur_atvik');
    [['fellt', 'Ferðir felldar niður'], ['atvik', 'Atvik'], ['kvort', 'Kvartanir']].forEach(([k, lbl]) => {
      const nu = summa(s.nu, k), ad = summa(s.adur, k);
      kpi(box, lbl + ' á skólaárinu', s.nu.length ? F.int(nu) : '—', s.nu.length && s.adur.length ? [breyting(signInt(nu - ad), k === 'fellt' ? null : nu <= ad), ' miðað við sömu mánuði í fyrra'] : null);
    });
    const ak = REIKNAD.akstur_atvik(sy);
    kpi(box, 'Atvik og kvartanir á mánuði', F.dec(ak), null, mA ? 'Viðmið ' + vidmidTexti(mA) : null, toneMk(mA, ak));
    graf('c-akstur', {
      type: 'bar',
      data: { labels: manadarMerki, datasets: [
        { label: 'Ferðir felldar niður', data: manadarRod(D.akstur, sy, 'fellt'), backgroundColor: L.grey, borderRadius: 3 },
        { label: 'Atvik', data: manadarRod(D.akstur, sy, 'atvik'), backgroundColor: L.c3, borderRadius: 3 },
        { label: 'Kvartanir', data: manadarRod(D.akstur, sy, 'kvort'), backgroundColor: L.bad, borderRadius: 3 },
      ] },
      options: { scales: { x: { stacked: true }, y: { stacked: true, beginAtZero: true, ticks: { precision: 0 } } },
        plugins: { tooltip: Object.assign(tipLabel(F.int), { callbacks: Object.assign(tipLabel(F.int).callbacks, manadarTip.callbacks) }) } },
    }, s.nu.length > 0);
  }

  /* ---------- Fjárhagur ---------- */
  function teiknaFjarhagur(sy) {
    const L = LITIR();
    const fy = fjarhagsar(sy); const r = D.fjarhagur.filter((x) => x.ar === fy);
    $('#fj-merki').textContent = fy ? `${fy}${r[0] && r[0].timabil ? ' · ' + r[0].timabil : ''}` : '';
    const box = $('#kpis-fjarhagur'); box.textContent = '';
    const a = summa(r, 'atil'), raun = summa(r, 'rtil'), frav = a ? raun / a - 1 : null;
    kpi(box, 'Áætlun ársins', r.length ? F.mkr(summa(r, 'aars')) : '—', 'millj. kr.');
    kpi(box, 'Raun til dags', r.length ? F.mkr(raun) : '—', r.length ? `af ${F.mkr(a)} millj. kr. áætlun til dags` : null);
    kpi(box, 'Frávik frá áætlun', signPct(frav), r.length ? F.sign(raun - a, F.mkr) + ' millj. kr.' : null, null, frav == null ? null : frav <= 0.005 ? 'good' : 'bad');

    const tb = $('#fj-tafla tbody'), tf = $('#fj-tafla tfoot'); tb.textContent = ''; tf.textContent = '';
    const lina = (parent, nafn, aars, atil, rtil) => {
      const tr = el('tr'); const fv = atil ? rtil / atil - 1 : null;
      tr.appendChild(el('td', null, nafn));
      tr.appendChild(el('td', 'num hide-sm', F.mkr(aars)));
      tr.appendChild(el('td', 'num', F.mkr(atil)));
      tr.appendChild(el('td', 'num', F.mkr(rtil)));
      const td = el('td', 'num', signPct(fv)); if (fv != null && Math.abs(fv) >= 0.005) td.classList.add(fv > 0 ? 'bad' : 'good'); tr.appendChild(td);
      parent.appendChild(tr);
    };
    r.forEach((x) => lina(tb, x.flokkur, x.aars, x.atil, x.rtil));
    if (r.length) lina(tf, 'Alls', summa(r, 'aars'), a, raun);
    graf('c-fjarhagur', {
      type: 'bar',
      data: { labels: r.map((x) => x.flokkur), datasets: [
        { label: 'Áætlun til dags', data: r.map((x) => x.atil), backgroundColor: L.grey, borderRadius: 3 },
        { label: 'Raun til dags', data: r.map((x) => x.rtil), backgroundColor: r.map((x) => (x.rtil > x.atil * 1.005 ? L.bad : L.c1)), borderRadius: 3 },
      ] },
      options: { scales: { y: { beginAtZero: true } }, plugins: { tooltip: tipLabel((v) => F.mkr(v) + ' millj. kr.') } },
    }, r.length > 0, 'Engar fjárhagstölur skráðar.');
  }

  /* ------------------------------------------------------------------ */
  /* Stýring                                                             */
  /* ------------------------------------------------------------------ */
  const TEIKNA = { yfirlit: teiknaYfirlit, nemendur: teiknaNemendur, maeting: teiknaMaeting, kennsla: teiknaKennsla,
    starfsfolk: teiknaStarfsfolk, samskipti: teiknaSamskipti, oryggi: teiknaOryggi, akstur: teiknaAkstur, fjarhagur: teiknaFjarhagur };

  function teiknaAllt() {
    if (!D || !st.ar) return;
    document.querySelectorAll('.ar-merki').forEach((n) => { n.textContent = 'skólaárið ' + arMerki(st.ar); });
    Object.keys(TEIKNA).forEach((k) => {
      try { TEIKNA[k](st.ar); } catch (e) { console.error(k, e); villa(`Ekki tókst að teikna liðinn „${k}“: ${e.message}`); }
    });
  }

  function syndaFlipa(tab) {
    if (!TEIKNA[tab]) tab = 'yfirlit';
    st.tab = tab;
    document.querySelectorAll('.tab').forEach((s) => s.classList.toggle('active', s.id === 't-' + tab));
    document.querySelectorAll('#flipar a').forEach((a) => {
      if (a.dataset.tab === tab) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    Object.values(grofin).forEach((g) => g.resize());
  }

  function villa(t) { const v = $('#villur'); v.hidden = false; v.textContent = t; }

  function hlada(buf, uppruni) {
    try {
      D = lesaGogn(buf);
    } catch (e) {
      console.error(e);
      villa('Ekki tókst að lesa skjalið. Er þetta rétt Excel-skjal? (' + e.message + ')');
      return;
    }
    $('#villur').hidden = true;
    if (vantar.length) villa('Þessir flipar fundust ekki í skjalinu og liðir þeirra verða tómir: ' + vantar.join(', ') + '.');
    $('#byrjun').hidden = true;
    const skoli = D.stillingar['skóli'] || 'Mælaborð';
    $('#skoli').textContent = skoli;
    $('#undirtitill').textContent = D.stillingar['undirtitill'] || '';
    document.title = skoli + ' – mælaborð';
    $('#daemi-borði').hidden = norm(D.stillingar['dæmagögn']) !== 'já';

    const allirArgangar = new Set();
    ['nemendur', 'maeting', 'forfoll', 'starfsfolk', 'samskipti', 'oryggi', 'akstur', 'stundir', 'innramat', 'skraningar', 'arsm']
      .forEach((k) => D[k].forEach((r) => r.sy && allirArgangar.add(r.sy)));
    const ar = Array.from(allirArgangar).sort();
    const virk = new Set(D.maeting.concat(D.nemendur, D.forfoll).map((r) => r.sy));
    const sjalfgefid = ar.filter((y) => virk.has(y)).pop() || ar[ar.length - 1];
    const val = $('#ar-val'); val.textContent = '';
    ar.slice().reverse().forEach((y) => { const o = el('option', null, arMerki(y)); o.value = y; val.appendChild(o); });
    if (!ar.includes(st.ar)) st.ar = sjalfgefid;
    val.value = st.ar;

    const t = uppruni.breytt ? new Date(uppruni.breytt) : null;
    $('#heimild').textContent = `Gögn: ${uppruni.nafn}${uppruni.stadbundid ? ' (opnað af þessari tölvu)' : ''}` +
      (t && !isNaN(t) ? ` · síðast breytt ${t.toLocaleDateString('is-IS', { day: 'numeric', month: 'long', year: 'numeric' })}` : '');

    teiknaAllt();
    syndaFlipa((location.hash || '#yfirlit').slice(1));
  }

  function opnaSkra(file) {
    if (!file) return;
    const fr = new FileReader();
    fr.onload = () => hlada(new Uint8Array(fr.result), { nafn: file.name, breytt: file.lastModified, stadbundid: true });
    fr.onerror = () => villa('Ekki tókst að opna skjalið.');
    fr.readAsArrayBuffer(file);
  }

  function synaByrjun() {
    $('#byrjun').hidden = false;
    document.querySelectorAll('.tab').forEach((s) => s.classList.remove('active'));
  }

  function prentaUndirbuning() {
    document.body.classList.add('prentun');
    document.querySelectorAll('.tab').forEach((s) => s.classList.add('active'));
    Object.values(grofin).forEach((g) => g.resize());
  }
  function prentaLokid() {
    document.body.classList.remove('prentun');
    syndaFlipa(st.tab);
  }

  async function byrja() {
    const inntak = $('#skra-inntak');
    $('#opna-btn').addEventListener('click', () => inntak.click());
    inntak.addEventListener('change', () => { opnaSkra(inntak.files[0]); inntak.value = ''; });
    const drop = $('#drop');
    drop.addEventListener('click', () => inntak.click());
    drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); inntak.click(); } });
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    document.addEventListener('drop', (e) => { const f = e.dataTransfer && e.dataTransfer.files[0]; if (f) opnaSkra(f); });
    $('#ar-val').addEventListener('change', (e) => { st.ar = e.target.value; teiknaAllt(); syndaFlipa(st.tab); });
    window.addEventListener('hashchange', () => syndaFlipa(location.hash.slice(1)));
    $('#prenta-btn').addEventListener('click', () => {
      prentaUndirbuning();
      requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(() => window.print(), 150)));
    });
    window.addEventListener('beforeprint', prentaUndirbuning);
    window.addEventListener('afterprint', prentaLokid);
    if (window.matchMedia) {
      const mq = window.matchMedia('(prefers-color-scheme: dark)');
      if (mq.addEventListener) mq.addEventListener('change', () => { teiknaAllt(); syndaFlipa(st.tab); });
    }

    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
    Chart.defaults.color = litur('--muted');

    if (typeof XLSX === 'undefined' || typeof Chart === 'undefined') {
      villa('Forritasöfnin í möppunni lib/ fundust ekki. Gakktu úr skugga um að allar skrár síðunnar hafi verið settar inn.');
      return;
    }
    if (CFG.gagnaskra && location.protocol !== 'file:') {
      try {
        const res = await fetch(CFG.gagnaskra + (CFG.gagnaskra.includes('?') ? '&' : '?') + 't=' + Date.now(), { cache: 'no-store' });
        if (!res.ok) throw new Error(res.status);
        const buf = new Uint8Array(await res.arrayBuffer());
        hlada(buf, { nafn: CFG.gagnaskra.split('/').pop(), breytt: res.headers.get('Last-Modified') });
        return;
      } catch (e) { /* fellur niður í að velja skjal */ }
    }
    synaByrjun();
  }

  document.addEventListener('DOMContentLoaded', byrja);
})();
