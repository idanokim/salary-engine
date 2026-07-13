/*
 * engine.js — the salary validation engine, ported to JavaScript so the whole
 * check can run **in the browser**. A גולמי file is parsed and validated locally
 * and never uploaded, which sidesteps serverless request-body limits (a 12 MB
 * file would otherwise be rejected). This is a faithful port of calculate() /
 * load_golmi() in main.py — keep the two in sync.
 *
 * UMD: exposes `SalaryEngine` as a browser global and as a CommonJS module.
 */
(function (global) {
  'use strict';

  const MATCH_THRESHOLD = 1.0;
  const DEFAULT_TRACK = 1;
  const CODE_COMBINED_BASE = 10002; // שכר משולב
  const CODE_YESOD = 1;             // יסוד משולב
  const CODE_VETEK_TOSEFET = 2;     // תוספת ותק
  const BASE_CODES = new Set([CODE_COMBINED_BASE, CODE_YESOD, CODE_VETEK_TOSEFET]);
  // The גולמי "ותק לחישוב שכר" column is rounded to the nearest quarter-year,
  // so a correct base recomputed off it can drift a shekel or two — validate the
  // base against any seniority inside the ±0.125-yr rounding window.
  const SENIORITY_ROUND = 0.125;
  const STATUS = { VALID: 'valid', INVALID: 'invalid', NO_BASE: 'no_base', MULTI: 'multi_period' };

  function round(x, n) {
    const f = Math.pow(10, n);
    return Math.round((x + Number.EPSILON) * f) / f;
  }
  const round2 = (x) => round(x, 2);

  // Turn the raw lookups.json into fast lookup structures.
  function prepLookups(raw) {
    const labelToBase = {};
    for (const k in raw.darga) labelToBase[String(k).trim()] = Number(raw.darga[k]);
    const vetek = {};
    for (const t in raw.vetek) {
      const tn = parseInt(t, 10);
      const pairs = Object.keys(raw.vetek[t])
        .map((y) => [parseFloat(y), Number(raw.vetek[t][y])])
        .sort((a, b) => a[0] - b[0]);
      vetek[tn] = { pairs, map: new Map(pairs.map((p) => [p[0], p[1]])) };
    }
    const trackMax = {};
    for (const t in (raw.track_max || {})) trackMax[parseInt(t, 10)] = Number(raw.track_max[t]);
    const tracks = {};
    for (const t in (raw.tracks || {})) tracks[parseInt(t, 10)] = String(raw.tracks[t]);
    return { labelToBase, vetek, trackMax, tracks };
  }

  // Canonicalize a grade label to '<number>+' form. Some dumps write the '+'
  // as a prefix ('+17') instead of a suffix ('17+'); move it to the end.
  function normalizeGradeLabel(darga) {
    if (darga === null || darga === undefined) return null;
    let s = String(darga).trim();
    if (s.startsWith('+') && s.slice(1).trim()) s = s.slice(1).trim() + '+';
    return s;
  }

  function getGradeBase(lk, darga) {
    if (darga === null || darga === undefined || darga === '') return null;
    const v = lk.labelToBase[normalizeGradeLabel(darga)];
    return v === undefined ? null : v;
  }

  function getVatekMultiplier(lk, vatek, track) {
    track = parseInt(track, 10) || DEFAULT_TRACK;
    const t = lk.vetek[track] || lk.vetek[DEFAULT_TRACK];
    if (!t) return null;
    vatek = parseFloat(vatek);
    const cap = lk.trackMax[track];
    if (cap !== undefined) vatek = Math.min(vatek, cap);
    if (t.map.has(vatek)) return t.map.get(vatek);
    const keys = t.pairs;
    let lo = null, hi = null;
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] < vatek) lo = keys[i];
      else { hi = keys[i]; break; }
    }
    if (!lo) return keys[0][1];
    if (!hi) return keys[keys.length - 1][1];
    return lo[1] + (hi[1] - lo[1]) * (vatek - lo[0]) / (hi[0] - lo[0]);
  }

  // --- Component rules (החוקה) -----------------------------------------------
  // component_rules.json (served at /api/rules): percentage components carry the
  // exact base composition + every official rate; manual (ידני) components are
  // accepted as reported. Trust is self-calibrating per file — see trustedRuleCodes.
  const TRUST_MIN_MATCH = 0.97, TRUST_MIN_N = 20;

  function prepRules(raw) {
    const rules = {};
    for (const k in raw) rules[parseInt(k, 10)] = raw[k];
    return rules;
  }

  function checkWorkerComponents(rows, jobPct, rules) {
    const amounts = new Map();
    for (const r of rows) {
      const code = Number(r.comp_code);
      if (!Number.isNaN(code)) amounts.set(code, (amounts.get(code) || 0) + (r.amount || 0));
    }
    const checks = {};
    const jp = jobPct || 1.0;
    for (const k in rules) {
      const rule = rules[k];
      if (rule.type !== 'percent' && rule.type !== 'shekel') continue;
      let slip = 0;
      for (const c of rule.codes) slip += (amounts.get(c) || 0);
      if (Math.abs(slip) < 0.01) continue;
      let expected;
      if (rule.type === 'percent') {
        let base = 0;
        for (const c of rule.base_codes) base += (amounts.get(c) || 0);
        base += (rule.base_const || 0) * jp;
        if (base <= 0) continue;
        let best = rule.rates[0];
        for (const r of rule.rates) if (Math.abs(base * r - slip) < Math.abs(base * best - slip)) best = r;
        expected = round2(base * best);
      } else {
        // shekel: one of a fixed set of flat amounts × job% (e.g. גמול מינהל
        // 4983 ∈ {105,210,315}). The closest admissible amount is the standard.
        let best = rule.amounts[0];
        for (const a of rule.amounts) if (Math.abs(a * jp - slip) < Math.abs(best * jp - slip)) best = a;
        expected = round2(best * jp);
      }
      checks[k] = {
        slip: round2(slip), expected, diff: round2(expected - slip),
        ok: Math.abs(expected - slip) <= MATCH_THRESHOLD, name: rule.name,
      };
    }
    return checks;
  }

  // גמולי השתלמות — population-calibrated checks (mirror of
  // check_gmul_population in main.py; keep in sync).
  const GMUL_A = [647, 667, 4268], GMUL_B = [897, 4269];
  const GMUL_NOTE = 'חריגה מהערך התקני לקבוצת הדרגה — חשד להפרשי רטרו בתוך הרכיב';

  function checkGmulPopulation(results) {
    const aVals = new Map(), bGroups = new Map(), per = new Map();
    for (let i = 0; i < results.length; i++) {
      const r = results[i];
      if (r.status !== STATUS.VALID && r.status !== STATUS.INVALID) continue;
      const amt = new Map();
      for (const c of r.components) amt.set(c.code, (amt.get(c.code) || 0) + (c.expected || 0));
      const a = GMUL_A.reduce((s, c) => s + (amt.get(c) || 0), 0);
      const b = GMUL_B.reduce((s, c) => s + (amt.get(c) || 0), 0);
      let base = 0;
      for (const c of BASE_CODES) base += (amt.get(c) || 0);
      const job = r.job_pct || 1.0;
      per.set(i, [a, b, base, job, r.kod_darga]);
      if (a > 0.01) {
        const k = round2(a / job);
        aVals.set(k, (aVals.get(k) || 0) + 1);
      }
      if (b > 0.01) {
        const g = bGroups.get(r.kod_darga) || new Map();
        const k = (base > 1 && Math.abs(b - 0.09 * base) <= MATCH_THRESHOLD) ? '9%' : round2(b / job);
        g.set(k, (g.get(k) || 0) + 1);
        bGroups.set(r.kod_darga, g);
      }
    }
    const top = (m) => { let bk = null, bn = 0, t = 0; for (const [k, n] of m) { t += n; if (n > bn) { bn = n; bk = k; } } return [bk, bn, t]; };
    let aMode = null;
    { const [k, n, t] = top(aVals); if (t >= TRUST_MIN_N && n / t >= TRUST_MIN_MATCH) aMode = k; }
    const bDom = new Map();
    for (const [kod, g] of bGroups) {
      const [k, n, t] = top(g);
      if (t >= 5 && n / t >= 0.90) bDom.set(kod, k);
    }
    const out = new Map();
    for (const [i, [a, b, base, job, kod]] of per) {
      const flags = {};
      if (aMode !== null && a > 0.01) {
        const exp = round2(aMode * job);
        if (Math.abs(a - exp) > MATCH_THRESHOLD) {
          flags[667] = { slip: round2(a), expected: exp, diff: round2(exp - a),
                         ok: false, name: "גמול השתלמות א'", note: GMUL_NOTE };
        }
      }
      const dom = bDom.get(kod);
      if (dom !== undefined && b > 0.01) {
        const exp = dom === '9%' ? round2(0.09 * base) : round2(dom * job);
        if (Math.abs(b - exp) > MATCH_THRESHOLD) {
          flags[897] = { slip: round2(b), expected: exp, diff: round2(exp - b),
                         ok: false, name: "גמול השתלמות ב'", note: GMUL_NOTE };
        }
      }
      if (Object.keys(flags).length) out.set(i, flags);
    }
    return out;
  }

  function trustedRuleCodes(allChecks, rules) {
    const per = new Map();
    for (const checks of allChecks) {
      for (const code in checks) {
        if (!per.has(code)) per.set(code, [0, 0]);
        const p = per.get(code);
        p[0]++; if (checks[code].ok) p[1]++;
      }
    }
    // "stable" rules (fixed era-proof rate, trivial base — ענ"א 16.2%/8%) are
    // trusted even below the per-file sample threshold.
    const stable = new Set();
    for (const k in (rules || {})) if (rules[k].stable) stable.add(String(k));
    const trusted = new Set();
    for (const [code, [n, ok]] of per) {
      if (n >= TRUST_MIN_N && ok / n >= TRUST_MIN_MATCH) trusted.add(code);
      else if (stable.has(code) && (n < TRUST_MIN_N || ok / n >= TRUST_MIN_MATCH)) trusted.add(code);
    }
    return trusted;
  }

  function baseWithinTolerance(lk, gradeBase, vatek, track, jobPct, slipBase) {
    if (gradeBase === null) return null;
    track = parseInt(track, 10) || DEFAULT_TRACK;
    const cap = lk.trackMax[track];
    // vetek arrives truncation-corrected (+0.005); widen the lower edge by the
    // same amount so the correction never newly penalizes.
    let vlo = vatek - SENIORITY_ROUND - 0.005, vhi = vatek + SENIORITY_ROUND;
    if (cap !== undefined) { vlo = Math.min(vlo, cap); vhi = Math.min(vhi, cap); }
    const mlo = getVatekMultiplier(lk, vlo, track);
    const mhi = getVatekMultiplier(lk, vhi, track);
    if (mlo === null || mhi === null) return null;
    const j = jobPct || 1.0;
    const lo = Math.min(mlo, mhi) * gradeBase * j;
    const hi = Math.max(mlo, mhi) * gradeBase * j;
    return (lo - MATCH_THRESHOLD) <= slipBase && slipBase <= (hi + MATCH_THRESHOLD);
  }

  // Map a גולמי column header to a canonical field name (port of _classify_header).
  function classifyHeader(h) {
    h = String(h).trim();
    if (h.includes('מסד') || h.includes('מסב') || h.includes('מזהה') || (h.includes('מספר') && h.includes('עובד'))) return 'worker_id';
    if (h.includes('קוד משרד') || h.includes('קוד גוף') || (h.includes('קוד') && h.includes('משרד/גוף'))) return 'ministry_code';
    if (h.includes('שם משרד') || h.includes('שם גוף') || h === 'משרד/גוף' || h === 'משרד' || h === 'גוף') return 'ministry_name';
    if (h.includes('דרוג')) return 'droog';
    if (h.includes('חלקיות')) return 'job_pct';
    if (h.includes('קוד דרגה')) return 'kod_darga';
    if (h.includes('קוד רכיב')) return 'comp_code';
    if (h.includes('רכיב')) return 'comp_name';
    if (h.includes('ותק')) return 'vatek';
    if (h === 'דרגה' || (h.includes('דרגה') && !h.includes('קוד'))) return 'darga_label';
    if (h.includes('פנסיו')) return 'pensionable';
    if (h.includes('סכום') || h.includes('סך')) return 'amount';
    return null;
  }

  // rows: array of arrays (sheet values). Returns Map<worker_id, component rows>.
  function loadGolmi(rows) {
    let headerMap = null, headerIdx = null, best = 0;
    const scan = rows.slice(0, 8);
    for (let i = 0; i < scan.length; i++) {
      const m = {};
      const row = scan[i] || [];
      for (let ci = 0; ci < row.length; ci++) {
        const cell = row[ci];
        if (cell === null || cell === undefined || cell === '') continue;
        const f = classifyHeader(cell);
        if (f && !(f in m)) m[f] = ci;
      }
      const score = Object.keys(m).length;
      if (score > best && 'worker_id' in m && 'comp_code' in m && 'amount' in m) {
        best = score; headerMap = m; headerIdx = i;
      }
    }
    if (!headerMap) throw new Error('לא נמצאה שורת כותרת מוכרת בגיליון הגולמי.');
    const get = (row, key) => {
      const idx = headerMap[key];
      return (idx !== undefined && idx < row.length) ? row[idx] : null;
    };
    const workers = new Map();
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const wid = get(row, 'worker_id');
      if (wid === null || wid === undefined || wid === '') continue;
      if (!workers.has(wid)) workers.set(wid, []);
      const amt = get(row, 'amount');
      workers.get(wid).push({
        ministry_code: get(row, 'ministry_code'), ministry_name: get(row, 'ministry_name'),
        droog: get(row, 'droog'), job_pct: get(row, 'job_pct'),
        kod_darga: get(row, 'kod_darga'), darga_label: get(row, 'darga_label'), vatek: get(row, 'vatek'),
        comp_code: get(row, 'comp_code'), comp_name: get(row, 'comp_name'),
        pensionable: get(row, 'pensionable'), amount: Number(amt) || 0,
      });
    }
    return workers;
  }

  // --- Plus-grade resolution (port of resolve_plus_grades) -------------------
  // Older מנהלי dumps drop the '+' from the דרגה label (a '18+' worker shows
  // '18') while קוד דרגה still distinguishes the plus grade. Self-calibrating
  // per file: each (kod_darga, label) population votes plain-vs-plus by which
  // grade reconstructs its slips; a code is remapped only on a decisive vote.
  const PLUS_MIN_VOTES = 5, PLUS_MIN_SHARE = 0.90;

  function resolvePlusGrades(lk, workers) {
    const votes = new Map(); // "kod|label" -> [plainFits, plusFits]
    for (const [, rows] of workers) {
      const first = rows[0];
      const label = String(first.darga_label ?? '').trim();
      if (!label || label.endsWith('+')) continue;
      const plainBase = getGradeBase(lk, label);
      const plusBase = getGradeBase(lk, label + '+');
      if (plainBase === null || plusBase === null) continue;
      // Only active single-period slips can vote.
      let slipBase = 0, cYesod = 0, cComb = 0;
      for (const r of rows) {
        const code = Number(r.comp_code);
        if (BASE_CODES.has(code)) slipBase += (r.amount || 0);
        if (code === CODE_YESOD) cYesod++;
        if (code === CODE_COMBINED_BASE) cComb++;
      }
      if (slipBase <= MATCH_THRESHOLD || Math.max(cYesod, cComb) > 1) continue;
      const track = parseInt(first.droog, 10) || DEFAULT_TRACK;
      const vatek = parseFloat(first.vatek) || 0;
      const jobPct = (parseFloat(first.job_pct) || 0) || 1.0;
      const plainOk = baseWithinTolerance(lk, plainBase, vatek, track, jobPct, slipBase);
      const plusOk = baseWithinTolerance(lk, plusBase, vatek, track, jobPct, slipBase);
      const key = `${first.kod_darga}|${label}`;
      if (plainOk && !plusOk) {
        if (!votes.has(key)) votes.set(key, [0, 0]);
        votes.get(key)[0]++;
      } else if (plusOk && !plainOk) {
        if (!votes.has(key)) votes.set(key, [0, 0]);
        votes.get(key)[1]++;
      }
    }
    const remap = new Map();
    for (const [key, [plainFits, plusFits]] of votes) {
      const decided = plainFits + plusFits;
      if (decided >= PLUS_MIN_VOTES && plusFits / decided >= PLUS_MIN_SHARE) {
        remap.set(key, key.slice(key.indexOf('|') + 1) + '+');
      }
    }
    return remap;
  }

  // Undo the גולמי file's 2-decimal truncation of seniority: the payroll grid
  // is in eighths (14.125) but the export truncates to 14.12 — a value that is
  // not a whole quarter is a truncated eighth; restore it (+0.005).
  function normalizeVatek(v) {
    v = parseFloat(v) || 0;
    if (round(v * 4, 6) % 1 !== 0) return round(v + 0.005, 3);
    return v;
  }

  function calculate(lk, rows, workerId, dargaOverride) {
    const first = rows[0];
    const track = parseInt(first.droog, 10) || DEFAULT_TRACK;
    const darga = (dargaOverride !== undefined && dargaOverride !== null)
      ? dargaOverride : normalizeGradeLabel(first.darga_label);
    const vatek = normalizeVatek(first.vatek);
    const jobPct = (parseFloat(first.job_pct) || 0) || 1.0;
    const gradeBase = getGradeBase(lk, darga);
    const vatekMult = getVatekMultiplier(lk, vatek, track);

    let rawBaseSum = 0, cYesod = 0, cComb = 0;
    for (const r of rows) {
      const code = Number(r.comp_code);
      if (BASE_CODES.has(code)) rawBaseSum += (r.amount || 0);
      if (code === CODE_YESOD) cYesod++;
      if (code === CODE_COMBINED_BASE) cComb++;
    }
    const primary = Math.max(cYesod, cComb);
    let status = null;
    if (rawBaseSum <= MATCH_THRESHOLD) status = STATUS.NO_BASE;
    else if (primary > 1) status = STATUS.MULTI;
    const recompute = status === null;

    let total = 0;
    const comps = [];
    for (const r of rows) {
      const code = Number(r.comp_code);
      const raw = r.amount || 0;
      let amount = raw, calculated = false, diff = null, computed = null;
      if (recompute && gradeBase !== null && vatekMult !== null) {
        if (code === CODE_COMBINED_BASE) computed = round2(gradeBase * vatekMult * jobPct);
        else if (code === CODE_YESOD) computed = round2(gradeBase * jobPct);
        else if (code === CODE_VETEK_TOSEFET) computed = round2(gradeBase * (vatekMult - 1.0) * jobPct);
      }
      if (computed !== null) { diff = round(computed - raw, 4); amount = computed; calculated = true; }
      total += amount;
      comps.push({ code: code || 0, name: r.comp_name || '', amount, calculated, expected: raw, diff,
                   pensionable: (r.pensionable === 'כן') });
    }
    let expectedTotal = 0;
    for (const r of rows) expectedTotal += (r.amount || 0);
    let totalDiff = round(total - expectedTotal, 4);
    let totalMatch = Math.abs(totalDiff) <= MATCH_THRESHOLD;
    if (status === null) {
      const ok = baseWithinTolerance(lk, gradeBase, vatek, track, jobPct, rawBaseSum);
      if (ok !== null) totalMatch = ok;
      // An unverifiable active slip (unknown grade/track) is never תקין by default.
      if (gradeBase === null || vatekMult === null) totalMatch = false;
      status = totalMatch ? STATUS.VALID : STATUS.INVALID;
    }
    if (status === STATUS.NO_BASE || status === STATUS.MULTI) totalMatch = null;
    return {
      worker_id: workerId, ministry_code: first.ministry_code || 0,
      ministry_name: first.ministry_name || '', droog: track,
      kod_darga: first.kod_darga || 0, darga_label: darga || '',
      vatek, job_pct: jobPct, grade_base: gradeBase, vatek_mult: vatekMult,
      components: comps, total: round2(total), expected_total: round2(expectedTotal),
      total_diff: totalDiff, total_match: totalMatch, status,
    };
  }

  function runEngine(lk, workers, rules) {
    const results = [];
    const allChecks = [];
    // Pass A0 — resolve dropped-'+' grade labels from the file's own population.
    const plusRemap = resolvePlusGrades(lk, workers);
    // Pass A — base validation + חוקה component checks per worker.
    for (const [wid, rows] of workers) {
      const first = rows[0];
      const key = `${first.kod_darga}|${String(first.darga_label ?? '').trim()}`;
      const r = calculate(lk, rows, wid, plusRemap.get(key));
      r.comp_flags = {};
      const active = r.status === STATUS.VALID || r.status === STATUS.INVALID;
      const checks = (rules && active) ? checkWorkerComponents(rows, r.job_pct, rules) : {};
      allChecks.push(checks);
      results.push(r);
    }
    if (!rules) {
      for (const r of results) r.findings = diagnoseResult(lk, r);
      return results;
    }
    // Pass B — self-calibrate rule trust on this file, then attach flags.
    const trusted = trustedRuleCodes(allChecks, rules);
    const gmulFlags = checkGmulPopulation(results);
    for (let i = 0; i < results.length; i++) {
      const r = results[i], checks = allChecks[i];
      for (const code in checks) {
        if (trusted.has(code) && !checks[code].ok) r.comp_flags[code] = checks[code];
      }
      Object.assign(r.comp_flags, gmulFlags.get(i) || {});
      const flagged = Object.keys(r.comp_flags);
      if (flagged.length) {
        if (r.status === STATUS.VALID) { r.status = STATUS.INVALID; r.total_match = false; }
        // Fold the component corrections into the simulator total, so
        // "סכום מחושב" is the full corrected slip — base AND components.
        let compDiff = 0;
        for (const k of flagged) compDiff += r.comp_flags[k].diff;
        compDiff = round2(compDiff);
        if (compDiff) {
          r.total = round2(r.total + compDiff);
          r.total_diff = round((r.total_diff || 0) + compDiff, 4);
        }
      }
      r.findings = diagnoseResult(lk, r);
    }
    return results;
  }

  // --- דוח שגויים: itemized error diagnosis per non-valid slip ---------------
  // Mirrors diagnose_entry() in main.py — keep the two in sync. Each finding:
  // {category, code?, name?, slip?, expected?, diff?, note} where diff is
  // always בתלוש minus תקני.
  const BASE_NAMES = { 1: 'יסוד משולב', 10002: 'שכר משולב', 2: 'תוספת ותק' };
  const STATUS_HE = { valid: 'תקין', invalid: 'שגוי',
                      no_base: 'ללא שכר בסיס פעיל', multi_period: 'רטרו / רב-תקופתי' };

  function impliedVatek(lk, track, impliedMult) {
    const t = lk.vetek[parseInt(track, 10) || DEFAULT_TRACK] || lk.vetek[DEFAULT_TRACK];
    if (!t) return null;
    let best = null, bestGap = Infinity;
    for (const [v, m] of t.pairs) {
      const g = Math.abs(m - impliedMult);
      if (g < bestGap) { bestGap = g; best = v; }
    }
    return bestGap <= 0.01 ? best : null;
  }

  function diagnoseResult(lk, r) {
    const findings = [];
    if (r.status === STATUS.NO_BASE) {
      if (Math.abs(r.expected_total || 0) <= MATCH_THRESHOLD) {
        findings.push({ category: 'תלוש ריק', note: 'אין סכומים בתלוש (סה"כ ≈ 0)' });
      } else {
        findings.push({ category: 'ללא שכר בסיס פעיל',
          note: `אין רכיב בסיס (יסוד/משולב) פעיל; סה"כ שאר הרכיבים ${r.expected_total}` });
      }
      return findings;
    }
    if (r.status === STATUS.MULTI) {
      findings.push({ category: 'רטרו / רב-תקופתי',
        note: 'רכיב בסיס ראשי מופיע יותר מפעם אחת — תלוש מרובה תקופות שכר' });
      return findings;
    }
    if (r.status !== STATUS.INVALID) return findings;

    if (r.grade_base === null || r.grade_base === undefined) {
      findings.push({ category: 'דרגה לא מוכרת',
        note: `דרגה '${r.darga_label}' לא נמצאה בטבלת השכר (קוד דרגה ${r.kod_darga})` });
    }

    let baseSlip = 0, baseExp = 0;
    const baseComps = [];
    for (const c of r.components) {
      if (c.calculated && BASE_NAMES[c.code]) {
        baseSlip += (c.expected || 0); baseExp += c.amount;
        baseComps.push(c);
      }
    }
    const baseGap = round2(baseSlip - baseExp);
    if (Math.abs(baseGap) > MATCH_THRESHOLD && baseExp > 0) {
      const ratio = baseSlip / baseExp;
      const doubled = Math.abs(ratio - 2.0) <= 0.01;
      if (doubled) {
        findings.push({ category: 'בסיס כפול (2×)',
          slip: round2(baseSlip), expected: round2(baseExp), diff: baseGap,
          note: 'סכומי הבסיס בתלוש כפולים בדיוק מהתקן — חשד לתלוש רטרו / שתי תקופות מאוחדות' });
      }
      baseComps.sort((a, b) => a.code - b.code);
      for (const c of baseComps) {
        const gap = round2((c.expected || 0) - c.amount);
        if (Math.abs(gap) > MATCH_THRESHOLD) {
          findings.push({ category: `פער ב${BASE_NAMES[c.code]}`,
            code: c.code, name: BASE_NAMES[c.code],
            slip: round2(c.expected || 0), expected: round2(c.amount), diff: gap });
        }
      }
      if (r.grade_base && !doubled) {
        const implied = baseSlip / (r.grade_base * (r.job_pct || 1.0));
        const v = impliedVatek(lk, r.droog, implied);
        if (v !== null) {
          const yearsGap = round2(v - (r.vatek || 0));
          if (Math.abs(yearsGap) >= 0.5) {
            findings.push({ category: 'ותק משתמע שונה מהרשום',
              note: `הבסיס בתלוש תואם ותק של כ-${v} שנים (מקדם ${round(implied, 4)}), ` +
                    `לעומת ${r.vatek} הרשום — פער ${yearsGap} שנים; ` +
                    `ייתכן ותק-לתשלום שונה או שינוי דרגה במהלך החודש` });
          }
        }
      }
    }

    const flagKeys = Object.keys(r.comp_flags || {}).sort((a, b) => a - b);
    for (const k of flagKeys) {
      const chk = r.comp_flags[k];
      findings.push({ category: `פער ב${chk.name}`, code: parseInt(k, 10), name: chk.name,
        slip: chk.slip, expected: chk.expected, diff: round2(chk.slip - chk.expected),
        note: chk.note || 'לפי נוסחת החוקה (אחוז × סמלי הבסיס בתלוש)' });
    }

    if (!findings.length) {
      findings.push({ category: 'פער כולל',
        slip: r.expected_total, expected: r.total,
        diff: round2((r.expected_total || 0) - (r.total || 0)),
        note: 'הסכום הכולל אינו תואם את החישוב' });
    }
    return findings;
  }

  const REPORT_HEADERS = ['מספר עובד', 'קוד משרד', 'שם משרד', 'דרגה', 'ותק', 'חלקיות',
    'סטטוס', 'קטגוריית שגיאה', 'סמל', 'שם רכיב', 'בתלוש', 'תקני', 'הפרש', 'פירוט'];

  function reportRows(results) {
    const rows = [];
    for (const r of results) {
      if (r.status === STATUS.VALID) continue;
      for (const f of (r.findings || [])) {
        rows.push([r.worker_id, r.ministry_code, r.ministry_name, r.darga_label,
          r.vatek, r.job_pct, STATUS_HE[r.status] || r.status,
          f.category ?? null, f.code ?? null, f.name ?? null,
          f.slip ?? null, f.expected ?? null, f.diff ?? null, f.note ?? null]);
      }
    }
    return rows;
  }

  function accuracyReport(results, elapsedSec) {
    const total = results.length;
    let valid = 0, invalid = 0, no_base = 0, multi = 0;
    for (const r of results) {
      if (r.status === STATUS.VALID) valid++;
      else if (r.status === STATUS.INVALID) invalid++;
      else if (r.status === STATUS.NO_BASE) no_base++;
      else if (r.status === STATUS.MULTI) multi++;
    }
    const active = results.filter((r) => r.status === STATUS.VALID || r.status === STATUS.INVALID);
    const activeTotal = active.length;
    const acc = activeTotal ? round(valid / activeTotal * 100, 2) : 0;
    const diffs = active.map((r) => Math.abs(r.total_diff));
    const avgDiff = diffs.length ? round(diffs.reduce((a, b) => a + b, 0) / diffs.length, 4) : 0;
    const maxDiff = diffs.length ? round(Math.max.apply(null, diffs), 4) : 0;
    const byMin = new Map();
    for (const r of active) {
      const k = r.ministry_name || '';
      if (!byMin.has(k)) byMin.set(k, { ministry_name: k, workers: 0, matched: 0 });
      const o = byMin.get(k); o.workers++; if (r.status === STATUS.VALID) o.matched++;
    }
    const by_ministry = Array.from(byMin.values())
      .map((o) => ({ ministry_name: o.ministry_name, workers: o.workers, matched: o.matched,
                     accuracy_pct: round(o.matched / o.workers * 100, 2) }))
      .sort((a, b) => b.workers - a.workers).slice(0, 20);
    // r.total / r.total_diff already include the חוקה component corrections
    // (folded in by runEngine) — use them directly.
    const inv = active.filter((r) => r.status === STATUS.INVALID)
      .sort((a, b) => Math.abs(b.total_diff) - Math.abs(a.total_diff)).slice(0, 300);
    const mismatches = inv.map((r) => ({
      worker_id: r.worker_id, ministry_name: r.ministry_name, darga_label: r.darga_label,
      vatek: r.vatek, job_pct: r.job_pct, grade_base: r.grade_base, vatek_multiplier: r.vatek_mult,
      total_calculated: round2(r.total), total_expected: r.expected_total,
      total_diff: round2(r.total_diff),
      components: r.components.filter((c) => c.calculated).map((c) => ({
        code: c.code, name: c.name, slip: round2(c.expected || 0),
        computed: round2(c.amount || 0), diff: round2(c.diff || 0),
      })).concat(Object.keys(r.comp_flags || {}).map((k) => ({
        code: parseInt(k, 10), name: r.comp_flags[k].name, slip: r.comp_flags[k].slip,
        computed: r.comp_flags[k].expected, diff: r.comp_flags[k].diff,
      }))),
      findings: (r.findings || []).map((f) => ({
        category: f.category,
        text: f.note || (f.code !== undefined && f.slip !== undefined
          ? `בתלוש ${f.slip} · תקני ${f.expected} · הפרש ${f.diff}` : ''),
      })),
    }));
    return {
      total_workers: total, matched: valid, unmatched: invalid, no_base, multi_period: multi,
      active_total: activeTotal, active_accuracy_pct: acc, accuracy_pct: acc,
      match_threshold: MATCH_THRESHOLD, avg_diff: avgDiff, max_diff: maxDiff,
      by_ministry, mismatches, elapsed_sec: elapsedSec,
    };
  }

  const BATCH_COLUMNS = ['worker_id', 'ministry_code', 'ministry_name', 'droog', 'kod_darga',
    'darga_label', 'vatek', 'job_pct', 'grade_base', 'vatek_mult', 'total_calculated',
    'total_expected', 'total_diff', 'gmul_diff', 'total_match', 'status',
    'flagged_components', 'n_components', 'errors'];

  function batchRow(r) {
    const flags = r.comp_flags || {};
    const flagged = Object.keys(flags).sort()
      .map((k) => `${k} (${flags[k].name}): ${flags[k].slip} במקום ${flags[k].expected}`)
      .join('; ');
    // gap in גמולי השתלמות (slip minus standard), one column like the CLI report
    let gmulDiff = 0;
    for (const k of [667, 897]) if (flags[k]) gmulDiff += flags[k].slip - flags[k].expected;
    gmulDiff = gmulDiff ? round2(gmulDiff) : null;
    return [r.worker_id, r.ministry_code, r.ministry_name, r.droog, r.kod_darga, r.darga_label,
      r.vatek, r.job_pct, r.grade_base, r.vatek_mult, r.total, r.expected_total, r.total_diff,
      gmulDiff, r.total_match, r.status, flagged, r.components.length, ''];
  }

  function batchCSV(results) {
    const esc = (v) => {
      if (v === null || v === undefined) return '';
      const s = String(v);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    const lines = [BATCH_COLUMNS.join(',')];
    for (const r of results) lines.push(batchRow(r).map(esc).join(','));
    return '﻿' + lines.join('\n');
  }

  // Pivot data for the highlighted "גולמי מעודכן" export (the ExcelJS writing,
  // which needs the browser library, lives in the page).
  function buildPivot(results) {
    const codeNames = {};
    const allCodes = new Set();
    const rows = [];
    for (const r of results) {
      const slipByCode = new Map();
      for (const c of r.components) {
        const code = c.code;
        if (!code) continue;
        slipByCode.set(code, (slipByCode.get(code) || 0) + (c.expected || 0));
        if (!(code in codeNames)) codeNames[code] = c.name || String(code);
        allCodes.add(code);
      }
      const invalidCodes = new Map();
      if (r.status === STATUS.INVALID) {
        for (const c of r.components) {
          if (c.calculated && c.diff !== null && Math.abs(c.diff) > MATCH_THRESHOLD) {
            invalidCodes.set(c.code, { computed: c.amount, slip: c.expected, diff: (c.expected || 0) - c.amount });
          }
        }
      }
      // חוקה component flags: expected per the rulebook vs the slip amount.
      // (r.total already includes these corrections — folded in by runEngine.)
      const flags = r.comp_flags || {};
      for (const k in flags) {
        const code = parseInt(k, 10), chk = flags[k];
        if (!invalidCodes.has(code)) {
          invalidCodes.set(code, { computed: chk.expected, slip: chk.slip,
                                   diff: chk.slip - chk.expected, rulebook: true });
        }
      }
      rows.push({
        meta: [r.worker_id, r.ministry_code, r.ministry_name, r.job_pct, r.kod_darga, r.darga_label, r.vatek],
        slipByCode, slipTotal: r.expected_total,
        correctedTotal: round2(r.total),
        invalidCodes, totalInvalid: r.status === STATUS.INVALID,
      });
    }
    const codesSorted = Array.from(allCodes).sort((a, b) => a - b);
    return { codesSorted, codeNames, rows };
  }

  const api = {
    MATCH_THRESHOLD, STATUS, round2,
    prepLookups, prepRules, getGradeBase, getVatekMultiplier, baseWithinTolerance,
    checkWorkerComponents, trustedRuleCodes, resolvePlusGrades, normalizeGradeLabel,
    diagnoseResult, reportRows, REPORT_HEADERS,
    classifyHeader, loadGolmi, calculate, runEngine,
    accuracyReport, batchCSV, BATCH_COLUMNS, batchRow, buildPivot,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.SalaryEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
