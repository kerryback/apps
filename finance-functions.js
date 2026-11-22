/*
 * finance-functions.js — Excel finance and statistics functions that
 * HyperFormula 2.7 lacks, written to match Excel's results.
 *
 *   Cash flow:     IRR, XIRR
 *   Depreciation:  VDB
 *   Regression:    INTERCEPT, FORECAST, FORECAST.LINEAR
 *   Statistics:    PERCENTILE, PERCENTILE.INC, PERCENTILE.EXC, QUARTILE,
 *                  QUARTILE.INC, QUARTILE.EXC, RANK, RANK.EQ, RANK.AVG, KURT
 *   Bonds:         PRICE, YIELD, DURATION, MDURATION, ACCRINT, ACCRINTM,
 *                  COUPDAYS, COUPDAYBS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD,
 *                  DISC, INTRATE, RECEIVED, PRICEDISC, PRICEMAT, YIELDDISC,
 *                  YIELDMAT
 *
 * Load it after HyperFormula and Handsontable, and before building any
 * HyperFormula instance or `new Handsontable(...)`. The order matters:
 * handsontable.full.min.js carries its own copy of HyperFormula and replaces
 * window.HyperFormula with it, so functions registered before it loads are lost.
 *
 *   <script src=".../hyperformula.full.min.js"></script>
 *   <script src=".../handsontable.full.min.js"></script>
 *   <script src="finance-functions.js"></script>
 *
 * Pass the named expressions when building the engine, so TRUE and FALSE work
 * (Handsontable ignores formulas.namedExpressions for an engine you supply):
 *   const hf = HyperFormula.buildEmpty({ licenseKey: '...' },
 *                                      FinanceFunctions.namedExpressions);
 *
 * It registers itself on window.HyperFormula. In Node:
 *   require('./finance-functions.js').install(HyperFormula)
 *
 * A function that HyperFormula already has is never replaced, so a later
 * HyperFormula that adds, say, IRR keeps its own.
 */
(function (root) {
  'use strict';

  // ---------- dates (Excel serial numbers, 1900 system) ----------
  const EPOCH = Date.UTC(1899, 11, 30);
  const DAY = 86400000;
  function toYmd(s) {
    const d = new Date(EPOCH + Math.floor(s) * DAY);
    return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
  }
  function toSerial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY); }
  function isLeap(y) { return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0; }
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function isLastDay(s) { const [y, m, d] = toYmd(s); return d === daysInMonth(y, m); }
  function addMonths(s, n, endOfMonth) {
    const [y, m, d] = toYmd(s);
    const t = y * 12 + (m - 1) + n;
    const ny = Math.floor(t / 12), nm = t - ny * 12 + 1;
    const dim = daysInMonth(ny, nm);
    return toSerial(ny, nm, endOfMonth ? dim : Math.min(d, dim));
  }

  // 30/360 day counts: US (NASD, as in Excel's YEARFRAC basis 0) and European.
  function days360(s1, s2, european) {
    let [y1, m1, d1] = toYmd(s1), [y2, m2, d2] = toYmd(s2);
    if (european) {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
    } else {
      const feb1 = m1 === 2 && isLastDay(s1), feb2 = m2 === 2 && isLastDay(s2);
      if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
      else if (d1 === 31) d1 = 30;
      else if (d1 === 30 && d2 === 31) d2 = 30;
      else if (feb1 && feb2) { d1 = 30; d2 = 30; }
      else if (feb1) d1 = 30;
    }
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
  }

  function yearFrac(s1, s2, basis) {
    s1 = Math.floor(s1); s2 = Math.floor(s2);
    if (s1 > s2) [s1, s2] = [s2, s1];
    switch (basis) {
      case 0: return days360(s1, s2, false) / 360;
      case 2: return (s2 - s1) / 360;
      case 3: return (s2 - s1) / 365;
      case 4: return days360(s1, s2, true) / 360;
      case 1: {
        const [y1, m1, d1] = toYmd(s1), [y2, m2, d2] = toYmd(s2);
        const withinYear = y1 === y2 || (y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && d1 >= d2)));
        if (withinYear) {
          let denom = 365;
          if (y1 === y2 && isLeap(y1)) denom = 366;
          else {
            const feb29 = y => isLeap(y) && s1 <= toSerial(y, 2, 29) && toSerial(y, 2, 29) <= s2;
            if (feb29(y1) || feb29(y2)) denom = 366;
          }
          return (s2 - s1) / denom;
        }
        const avg = (toSerial(y2 + 1, 1, 1) - toSerial(y1, 1, 1)) / (y2 - y1 + 1);
        return (s2 - s1) / avg;
      }
    }
    throw 'NUM';
  }

  // ---------- coupon schedule ----------
  function couponInfo(settle, mat, freq) {
    const step = 12 / freq, eom = isLastDay(mat);
    for (let k = 1; ; k++) {
      const d = addMonths(mat, -k * step, eom);
      if (d <= settle) {
        return { pcd: d, ncd: k === 1 ? mat : addMonths(mat, -(k - 1) * step, eom), num: k };
      }
    }
  }
  function couponDays(settle, mat, freq, basis) {        // E, A, DSC
    const c = couponInfo(settle, mat, freq);
    let E, A, DSC;
    if (basis === 1) E = c.ncd - c.pcd;
    else if (basis === 3) E = 365 / freq;
    else E = 360 / freq;
    if (basis === 0) A = days360(c.pcd, settle, false);
    else if (basis === 4) A = days360(c.pcd, settle, true);
    else A = settle - c.pcd;
    DSC = (basis === 0 || basis === 4) ? E - A : c.ncd - settle;
    return { E, A, DSC, ...c };
  }

  // ---------- checks ----------
  function intOf(x) { return Math.trunc(x); }
  function checkBasis(b) { b = intOf(b); if (b < 0 || b > 4) throw 'NUM'; return b; }
  function checkFreq(f) { f = intOf(f); if (f !== 1 && f !== 2 && f !== 4) throw 'NUM'; return f; }
  function checkDates(settle, mat) {
    settle = intOf(settle); mat = intOf(mat);
    if (settle >= mat) throw 'NUM';
    return [settle, mat];
  }

  // ---------- root finding ----------
  function solve(f, guess, lo, hi) {
    let r = guess;
    for (let i = 0; i < 100; i++) {                  // Newton with a numeric slope
      const v = f(r), h = 1e-7 * Math.max(1, Math.abs(r));
      const d = (f(r + h) - v) / h;
      if (!isFinite(v) || !isFinite(d) || d === 0) break;
      const next = r - v / d;
      if (!isFinite(next) || next <= lo) break;
      if (Math.abs(next - r) < 1e-12) return next;
      r = next;
    }
    let a = lo, b = hi, fa = f(a), fb = f(b);         // bisection fallback
    if (!isFinite(fa) || !isFinite(fb) || fa * fb > 0) return null;
    for (let i = 0; i < 300; i++) {
      const m = (a + b) / 2, fm = f(m);
      if (fm === 0 || (b - a) / 2 < 1e-15) return m;
      if (fa * fm < 0) b = m; else { a = m; fa = fm; }
    }
    return (a + b) / 2;
  }

  // ---------- cash flow ----------
  function irr(flows, guess) {
    if (!flows.some(x => x > 0) || !flows.some(x => x < 0)) throw 'NUM';
    const npv = r => flows.reduce((s, x, t) => s + x / Math.pow(1 + r, t), 0);
    const r = solve(npv, guess, -0.999999999, 1e10);
    if (r === null) throw 'NUM';
    return r;
  }
  function xirr(values, dates, guess) {
    if (values.length !== dates.length || values.length < 2) throw 'NUM';
    if (!values.some(x => x > 0) || !values.some(x => x < 0)) throw 'NUM';
    const d0 = intOf(dates[0]);
    const ds = dates.map(d => { d = intOf(d); if (d < d0) throw 'NUM'; return d; });
    const f = r => values.reduce((s, v, i) => s + v / Math.pow(1 + r, (ds[i] - d0) / 365), 0);
    const r = solve(f, guess, -0.999999999, 1e10);
    if (r === null) throw 'NUM';
    return r;
  }

  // ---------- depreciation ----------
  function ddbPeriod(cost, salvage, life, period, factor) {
    let rate = factor / life, oldV;
    if (rate >= 1) { rate = 1; oldV = period === 1 ? cost : 0; }
    else oldV = cost * Math.pow(1 - rate, period - 1);
    const newV = cost * Math.pow(1 - rate, period);
    const dep = newV < salvage ? oldV - salvage : oldV - newV;
    return dep < 0 ? 0 : dep;
  }
  // Depreciation from time 0 to `period` (fractional periods prorated),
  // switching to straight line when that is larger.
  function vdbTo(cost, salvage, life, period, factor) {
    let total = 0, rest = cost - salvage, straight = false, sl = 0;
    const end = Math.ceil(period);
    for (let i = 1; i <= end; i++) {
      let term;
      if (!straight) {
        const ddb = ddbPeriod(cost, salvage, life, i, factor);
        sl = rest / (life - (i - 1));
        if (sl > ddb) { term = sl; straight = true; }
        else { term = ddb; rest -= ddb; }
      } else term = sl;
      if (i === end) term *= period + 1 - end;
      total += term;
    }
    return total;
  }
  function vdb(cost, salvage, life, start, end, factor, noSwitch) {
    if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life || factor <= 0) throw 'NUM';
    if (noSwitch) {
      const s = Math.floor(start), e = Math.ceil(end);
      let total = 0;
      for (let i = s + 1; i <= e; i++) {
        let term = ddbPeriod(cost, salvage, life, i, factor);
        if (i === s + 1) term *= Math.min(end, s + 1) - start;
        else if (i === e) term *= end + 1 - e;
        total += term;
      }
      return total;
    }
    return vdbTo(cost, salvage, life, end, factor) - vdbTo(cost, salvage, life, start, factor);
  }

  // ---------- regression and statistics ----------
  function pairs(ys, xs) {
    if (ys.length !== xs.length) throw 'NA';
    if (ys.length < 2) throw 'DIV0';
    const n = ys.length, mx = xs.reduce((a, b) => a + b) / n, my = ys.reduce((a, b) => a + b) / n;
    let sxy = 0, sxx = 0;
    for (let i = 0; i < n; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) ** 2; }
    if (sxx === 0) throw 'DIV0';
    const slope = sxy / sxx;
    return { slope, intercept: my - slope * mx };
  }
  function sorted(xs) { if (!xs.length) throw 'NUM'; return xs.slice().sort((a, b) => a - b); }
  function percentileInc(xs, k) {
    if (k < 0 || k > 1) throw 'NUM';
    const s = sorted(xs), h = (s.length - 1) * k, lo = Math.floor(h);
    return lo + 1 < s.length ? s[lo] + (h - lo) * (s[lo + 1] - s[lo]) : s[lo];
  }
  function percentileExc(xs, k) {
    const s = sorted(xs), n = s.length, h = (n + 1) * k;
    if (k <= 0 || k >= 1 || h < 1 || h > n) throw 'NUM';
    const lo = Math.floor(h);
    return lo < n ? s[lo - 1] + (h - lo) * (s[lo] - s[lo - 1]) : s[lo - 1];
  }
  function quartile(xs, q, exclusive) {
    q = intOf(q);
    if (exclusive ? (q < 1 || q > 3) : (q < 0 || q > 4)) throw 'NUM';
    return exclusive ? percentileExc(xs, q / 4) : percentileInc(xs, q / 4);
  }
  function rank(x, xs, order, average) {
    if (!xs.includes(x)) throw 'NA';
    const better = xs.filter(v => (order ? v < x : v > x)).length;
    const ties = xs.filter(v => v === x).length;
    return average ? better + (ties + 1) / 2 : better + 1;
  }
  function kurt(xs) {
    const n = xs.length;
    if (n < 4) throw 'DIV0';
    const m = xs.reduce((a, b) => a + b) / n;
    const s2 = xs.reduce((a, v) => a + (v - m) ** 2, 0) / (n - 1);
    if (s2 === 0) throw 'DIV0';
    const s4 = xs.reduce((a, v) => a + ((v - m) ** 2 / s2) ** 2, 0);
    return n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * s4 - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3));
  }

  // ---------- bonds ----------
  function price(settle, mat, rate, yld, red, freq, basis) {
    [settle, mat] = checkDates(settle, mat); freq = checkFreq(freq); basis = checkBasis(basis);
    if (rate < 0 || yld < 0 || red <= 0) throw 'NUM';
    return priceCore(settle, mat, rate, yld, red, freq, basis);
  }
  function priceCore(settle, mat, rate, yld, red, freq, basis) {
    const { E, A, DSC, num: N } = couponDays(settle, mat, freq, basis);
    // Excel documents one formula for PRICE, used even in the final coupon
    // period (YIELD, by contrast, switches to simple interest there).
    const c = 100 * rate / freq, y = yld / freq, x = DSC / E;
    let p = red / Math.pow(1 + y, N - 1 + x);
    for (let k = 1; k <= N; k++) p += c / Math.pow(1 + y, k - 1 + x);
    return p - c * A / E;
  }
  function yieldFn(settle, mat, rate, pr, red, freq, basis) {
    [settle, mat] = checkDates(settle, mat); freq = checkFreq(freq); basis = checkBasis(basis);
    if (rate < 0 || pr <= 0 || red <= 0) throw 'NUM';
    const { E, A, DSC, num: N } = couponDays(settle, mat, freq, basis);
    if (N === 1) {
      const c = rate / freq, a = pr / 100 + A / E * c;
      return (red / 100 + c - a) / a * freq * E / DSC;
    }
    const y = solve(r => priceCore(settle, mat, rate, r, red, freq, basis) - pr, 0.05, -freq + 1e-9, 100);
    if (y === null) throw 'NUM';
    return y;
  }
  function duration(settle, mat, coupon, yld, freq, basis) {
    [settle, mat] = checkDates(settle, mat); freq = checkFreq(freq); basis = checkBasis(basis);
    if (coupon < 0 || yld < 0) throw 'NUM';
    const { E, DSC, num: N } = couponDays(settle, mat, freq, basis);
    const c = 100 * coupon / freq, y = 1 + yld / freq, x = DSC / E;
    let pv = 0, tpv = 0;
    for (let k = 1; k <= N; k++) {
      const t = k - 1 + x, cf = c + (k === N ? 100 : 0), v = cf / Math.pow(y, t);
      pv += v; tpv += t * v;
    }
    return tpv / pv / freq;
  }
  function accrint(issue, first, settle, rate, par, freq, basis, calcMethod) {
    issue = intOf(issue); first = intOf(first); settle = intOf(settle);
    freq = checkFreq(freq); basis = checkBasis(basis);
    if (rate <= 0 || par <= 0 || issue >= settle) throw 'NUM';
    const start = !calcMethod && settle > first ? first : issue;
    return par * rate * yearFrac(start, settle, basis);
  }
  function accrintm(issue, settle, rate, par, basis) {
    issue = intOf(issue); settle = intOf(settle); basis = checkBasis(basis);
    if (rate <= 0 || par <= 0 || issue >= settle) throw 'NUM';
    return par * rate * yearDiff(issue, settle, basis);
  }
  // DISC and YIELDDISC measure the term with YEARFRAC. ACCRINTM, INTRATE,
  // RECEIVED and PRICEDISC, on basis 1, divide actual days by the days in the
  // starting year, as LibreOffice's Excel-compatible versions do. Over spans
  // shorter than a year the two agree.
  function yearDiff(s1, s2, basis) {
    if (basis !== 1) return yearFrac(s1, s2, basis);
    s1 = Math.floor(s1); s2 = Math.floor(s2);
    return (s2 - s1) / (isLeap(toYmd(s1)[0]) ? 366 : 365);
  }
  function discArgs(settle, mat, basis, diff) {
    [settle, mat] = checkDates(settle, mat); basis = checkBasis(basis);
    return diff ? yearDiff(settle, mat, basis) : yearFrac(settle, mat, basis);
  }
  function priceMat(settle, mat, issue, rate, yld, basis) {
    [settle, mat] = checkDates(settle, mat); basis = checkBasis(basis); issue = intOf(issue);
    if (rate < 0 || yld < 0) throw 'NUM';
    const dim = yearFrac(issue, mat, basis), a = yearFrac(issue, settle, basis), dsm = yearFrac(settle, mat, basis);
    return (100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100;
  }
  function yieldMat(settle, mat, issue, rate, pr, basis) {
    [settle, mat] = checkDates(settle, mat); basis = checkBasis(basis); issue = intOf(issue);
    if (rate < 0 || pr <= 0) throw 'NUM';
    const dim = yearFrac(issue, mat, basis), a = yearFrac(issue, settle, basis), dsm = yearFrac(settle, mat, basis);
    const b = pr / 100 + a * rate;
    return ((1 + dim * rate) - b) / b / dsm;
  }

  // ---------- registration with HyperFormula ----------
  function install(HF) {
    if (!HF || installedOn.has(HF.HyperFormula || HF)) return;
    // The browser build hangs everything on the HyperFormula class; the Node
    // module exports the class and the helpers side by side. Accept both.
    const M = HF;
    HF = HF.HyperFormula || HF;
    const T = M.FunctionArgumentType, CellError = M.CellError, ErrorType = M.ErrorType;
    const ERR = { NUM: ErrorType.NUM, NA: ErrorType.NA, DIV0: ErrorType.DIV_BY_ZERO, VALUE: ErrorType.VALUE };

    function num(v) {
      if (typeof v === 'number') return v;
      if (v && typeof v.val === 'number') return v.val;   // dates, percents, currency
      return null;
    }
    function numbers(range) {                              // numbers in a range; text and blanks skipped
      const out = [];
      for (const v of range.valuesFromTopLeftCorner()) {
        if (v instanceof CellError) throw v;
        if (typeof v === 'boolean') continue;
        const n = num(v);
        if (n !== null) out.push(n);
      }
      return out;
    }
    function numbersStrict(range) {                        // every cell must be a number (XIRR)
      return range.valuesFromTopLeftCorner().map(v => {
        if (v instanceof CellError) throw v;
        const n = num(v);
        if (n === null) throw 'VALUE';
        return n;
      });
    }
    function wrap(fn) {
      return function (...args) {
        try {
          const r = fn(...args.map(a => (a && typeof a.val === 'number' ? a.val : a)));
          if (typeof r === 'number' && !isFinite(r)) return new CellError(ErrorType.NUM);
          return r;
        } catch (e) {
          if (e instanceof CellError) return e;
          if (typeof e === 'string' && ERR[e]) return new CellError(ERR[e]);
          throw e;
        }
      };
    }

    const R = { argumentType: T.RANGE }, N = { argumentType: T.NUMBER };
    const opt = v => ({ argumentType: T.NUMBER, defaultValue: v });
    const optB = v => ({ argumentType: T.BOOLEAN, defaultValue: v });
    const defs = {
      'IRR':             [[R, opt(0.1)], (v, g) => irr(numbers(v), g)],
      'XIRR':            [[R, R, opt(0.1)], (v, d, g) => xirr(numbersStrict(v), numbersStrict(d), g)],
      'VDB':             [[N, N, N, N, N, opt(2), optB(false)], vdb],
      'INTERCEPT':       [[R, R], (y, x) => pairs(numbers(y), numbers(x)).intercept],
      'FORECAST':        [[N, R, R], (x, y, xs) => { const p = pairs(numbers(y), numbers(xs)); return p.intercept + p.slope * x; }],
      'FORECAST.LINEAR': [[N, R, R], (x, y, xs) => { const p = pairs(numbers(y), numbers(xs)); return p.intercept + p.slope * x; }],
      'PERCENTILE':      [[R, N], (a, k) => percentileInc(numbers(a), k)],
      'PERCENTILE.INC':  [[R, N], (a, k) => percentileInc(numbers(a), k)],
      'PERCENTILE.EXC':  [[R, N], (a, k) => percentileExc(numbers(a), k)],
      'QUARTILE':        [[R, N], (a, q) => quartile(numbers(a), q, false)],
      'QUARTILE.INC':    [[R, N], (a, q) => quartile(numbers(a), q, false)],
      'QUARTILE.EXC':    [[R, N], (a, q) => quartile(numbers(a), q, true)],
      'RANK':            [[N, R, opt(0)], (x, r, o) => rank(x, numbers(r), o, false)],
      'RANK.EQ':         [[N, R, opt(0)], (x, r, o) => rank(x, numbers(r), o, false)],
      'RANK.AVG':        [[N, R, opt(0)], (x, r, o) => rank(x, numbers(r), o, true)],
      'KURT':            [[R], a => kurt(numbers(a))],
      'PRICE':           [[N, N, N, N, N, N, opt(0)], price],
      'YIELD':           [[N, N, N, N, N, N, opt(0)], yieldFn],
      'DURATION':        [[N, N, N, N, N, opt(0)], duration],
      'MDURATION':       [[N, N, N, N, N, opt(0)], (s, m, c, y, f, b) => duration(s, m, c, y, f, b) / (1 + y / checkFreq(f))],
      'ACCRINT':         [[N, N, N, N, N, N, opt(0), optB(true)], accrint],
      'ACCRINTM':        [[N, N, N, opt(1000), opt(0)], accrintm],
      'COUPDAYS':        [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); return couponDays(s, m, checkFreq(f), checkBasis(b)).E; }],
      'COUPDAYBS':       [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); return couponDays(s, m, checkFreq(f), checkBasis(b)).A; }],
      'COUPDAYSNC':      [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); return couponDays(s, m, checkFreq(f), checkBasis(b)).DSC; }],
      'COUPNCD':         [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); checkBasis(b); return couponInfo(s, m, checkFreq(f)).ncd; }],
      'COUPPCD':         [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); checkBasis(b); return couponInfo(s, m, checkFreq(f)).pcd; }],
      'COUPNUM':         [[N, N, N, opt(0)], (s, m, f, b) => { [s, m] = checkDates(s, m); checkBasis(b); return couponInfo(s, m, checkFreq(f)).num; }],
      'DISC':            [[N, N, N, N, opt(0)], (s, m, pr, red, b) => { if (pr <= 0 || red <= 0) throw 'NUM'; return (red - pr) / red / discArgs(s, m, b); }],
      'INTRATE':         [[N, N, N, N, opt(0)], (s, m, inv, red, b) => { if (inv <= 0 || red <= 0) throw 'NUM'; return (red - inv) / inv / discArgs(s, m, b, true); }],
      'RECEIVED':        [[N, N, N, N, opt(0)], (s, m, inv, d, b) => { if (inv <= 0 || d <= 0) throw 'NUM'; const v = 1 - d * discArgs(s, m, b, true); if (v <= 0) throw 'NUM'; return inv / v; }],
      'PRICEDISC':       [[N, N, N, N, opt(0)], (s, m, d, red, b) => { if (d <= 0 || red <= 0) throw 'NUM'; return red * (1 - d * discArgs(s, m, b, true)); }],
      'YIELDDISC':       [[N, N, N, N, opt(0)], (s, m, pr, red, b) => { if (pr <= 0 || red <= 0) throw 'NUM'; return (red / pr - 1) / discArgs(s, m, b); }],
      'PRICEMAT':        [[N, N, N, N, N, opt(0)], priceMat],
      'YIELDMAT':        [[N, N, N, N, N, opt(0)], yieldMat],
    };

    const have = new Set(HF.getRegisteredFunctionNames('enGB'));
    class FinanceFunctions extends M.FunctionPlugin {}
    FinanceFunctions.implementedFunctions = {};
    const translations = {};
    // enGB (HyperFormula's default) may not be registered yet when this file
    // loads, so always include it.
    const langs = Array.from(new Set(['enGB', ...HF.getRegisteredLanguagesCodes()]));
    langs.forEach(l => { translations[l] = {}; });
    Object.entries(defs).forEach(([name, [params, fn]], i) => {
      if (have.has(name)) return;
      const method = 'fn' + i, impl = wrap(fn);
      FinanceFunctions.prototype[method] = function (ast, state) {
        return this.runFunction(ast.args, state, this.metadata(name), impl);
      };
      FinanceFunctions.implementedFunctions[name] = { method, parameters: params };
      langs.forEach(l => { translations[l][name] = name; });
    });
    HF.registerFunctionPlugin(FinanceFunctions, translations);
    installedOn.add(HF);
  }
  const installedOn = new WeakSet();   // each copy of HyperFormula gets the functions once

  // HyperFormula reads a bare TRUE or FALSE as an unknown name (#NAME?), so
  // VLOOKUP(..., FALSE) fails. Pass these as the second argument of
  // HyperFormula.buildEmpty or buildFromArray to make them work as in Excel.
  const namedExpressions = [
    { name: 'TRUE', expression: '=TRUE()' },
    { name: 'FALSE', expression: '=FALSE()' },
  ];

  const api = { install, namedExpressions, yearFrac, irr, xirr, vdb, price };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (root) root.FinanceFunctions = api;
  if (root && root.HyperFormula) install(root.HyperFormula);
})(typeof window !== 'undefined' ? window : globalThis);
