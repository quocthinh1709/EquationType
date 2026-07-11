// ==========================================
// THUẬT TOÁN TOÁN HỌC VÀ LÀM TRÒN SỐ
// ==========================================

const extractRootsRecursively = (node: any, targetArray: number[]) => {
    if (!node) return;
    if (typeof node === 'number') { targetArray.push(node); return; }
    if (typeof node.numericValue === 'number') { targetArray.push(node.numericValue); return; }
    if (typeof node.value === 'number') { targetArray.push(node.value); return; }

    if (node.head === "Equal" && node.ops && node.ops.length === 2) {
      extractRootsRecursively(node.ops[1], targetArray);
      return;
    }
    if (node.ops && Array.isArray(node.ops)) {
      node.ops.forEach((op: any) => extractRootsRecursively(op, targetArray));
    }
};

const getNum = (exprNode: any) => {
    const val = exprNode.evaluate();
    if (val.numericValue !== undefined && val.numericValue !== null) return val.numericValue;
    const p = parseFloat(val.toString());
    return isNaN(p) ? NaN : p;
};

const toNiceFraction = (x: number, tol = 1e-6, maxDen = 50): [number, number] | null => {
    if (!isFinite(x)) return null;
    const sign = x < 0 ? -1 : 1;
    let ax = Math.abs(x);
    let h1 = 1, h2 = 0, k1 = 0, k2 = 1, b = ax;
    let iter = 0;
    do {
      const a = Math.floor(b);
      [h1, h2] = [a * h1 + h2, h1];
      [k1, k2] = [a * k1 + k2, k1];
      const rem = b - a;
      if (rem < 1e-12) break;
      b = 1 / rem;
      iter++;
    } while (Math.abs(ax - h1 / k1) > ax * tol && k1 < maxDen && iter < 20);
    if (k1 > maxDen || k1 === 0) return null;
    return [sign * h1, k1];
};

const extractSquareFactor = (k: number): [number, number] => {
    let m = 1;
    let rem = k;
    for (let i = 2; i * i <= rem; i++) {
      while (rem % (i * i) === 0) {
        rem /= i * i;
        m *= i;
      }
    }
    return [m, rem];
};

const toNiceSqrtV2 = (ax: number, tol = 1e-6, maxDen = 50): { p: number; q: number; n: number } | null => {
    if (ax < 1e-9) return null;
    const r = ax * ax;
    const frac = toNiceFraction(r, tol, maxDen);
    if (!frac) return null;

    const [pNum, qDen] = frac; 
    if (pNum <= 0) return null;

    const pq = pNum * qDen;
    const [m, n] = extractSquareFactor(pq);

    if (n === 1) return null; 

    return { p: m, q: qDen, n };
};

const buildKnownValues = (): Record<string, string> => {
    const table: Record<string, string> = {};
    const add = (val: number, latex: string) => {
      const key = val.toFixed(12);
      if (!(key in table)) table[key] = latex;
    };

    const plainSqrts = [2, 3, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 17, 18, 20];
    plainSqrts.forEach((n) => add(Math.sqrt(n), `\\sqrt{${n}}`));

    add(0, "0"); add(0.5, "\\frac{1}{2}"); add(Math.sqrt(2) / 2, "\\frac{\\sqrt{2}}{2}");
    add(Math.sqrt(3) / 2, "\\frac{\\sqrt{3}}{2}"); add(1, "1");
    add(Math.sqrt(3) / 3, "\\frac{\\sqrt{3}}{3}"); add(Math.sqrt(3), "\\sqrt{3}");

    const fracSqrts: [number, number, number][] = [
      [1, 2, 2], [1, 2, 3], [1, 2, 4], [1, 3, 2], [1, 3, 3], [1, 3, 4],
      [1, 5, 2], [1, 5, 5], [2, 2, 3], [2, 3, 3], [3, 2, 2], [3, 3, 2],
      [1, 6, 2], [1, 6, 3], [1, 7, 2], [1, 7, 7], [1, 10, 2], [1, 10, 5],
      [1, 15, 3], [1, 15, 5],
    ];
    fracSqrts.forEach(([p, n, q]) => {
      const val = (p * Math.sqrt(n)) / q;
      const sqrtPart = `\\sqrt{${n}}`;
      const numerator = p === 1 ? sqrtPart : `${p}${sqrtPart}`;
      add(val, `\\frac{${numerator}}{${q}}`);
    });

    const cubeRoots = [2, 3, 4, 5, 6, 7, 9, 10];
    cubeRoots.forEach((n) => add(Math.cbrt(n), `\\sqrt[3]{${n}}`));

    add((1 + Math.sqrt(5)) / 2, "\\frac{1+\\sqrt{5}}{2}");
    add((Math.sqrt(5) - 1) / 2, "\\frac{\\sqrt{5}-1}{2}");
    add(Math.PI, "\\pi"); add(Math.PI / 2, "\\frac{\\pi}{2}");
    add(Math.PI / 3, "\\frac{\\pi}{3}"); add(Math.PI / 4, "\\frac{\\pi}{4}");
    add(Math.PI / 6, "\\frac{\\pi}{6}"); add(Math.E, "e");

    return table;
};

const KNOWN_VALUES = buildKnownValues();

const lookupNice = (x: number): string | null => {
    if (!isFinite(x)) return null;
    const ax = Math.abs(x);
    const key = ax.toFixed(12);
    const val = KNOWN_VALUES[key];
    if (!val) return null;
    return x < 0 && ax !== 0 ? `-${val}` : val;
};

const niceNumber = (x: number): string => {
    if (!isFinite(x)) return x > 0 ? "+\\infty" : "-\\infty";
    const r = Math.round(x * 1e6) / 1e6;
    const known = lookupNice(r);
    if (known) return known;
    if (Math.abs(r - Math.round(r)) < 1e-6) return String(Math.round(r));

    const frac = toNiceFraction(r);
    if (frac) {
      const [p, q] = frac;
      const sign = p < 0 ? "-" : "";
      return `${sign}${Math.abs(p)}/${q}`;
    }

    const ax = Math.abs(r);
    const sq = toNiceSqrtV2(ax);
    if (sq) {
      const { p, q, n } = sq;
      const sign = r < 0 ? "-" : "";
      const sqrtPart = `\\sqrt{${n}}`;
      if (q === 1) {
        return p === 1 ? `${sign}${sqrtPart}` : `${sign}${p}${sqrtPart}`;
      } else {
        const numerator = p === 1 ? sqrtPart : `${p}${sqrtPart}`;
        return `${sign}\\frac{${numerator}}{${q}}`;
      }
    }
    return String(Math.round(x * 100) / 100);
};

const findRootsNumeric = (f: (x: number) => number, lo = -50, hi = 50, steps = 4000): number[] => {
    const roots: number[] = [];
    const dx = (hi - lo) / steps;
    let prevX = lo;
    let prevY = f(prevX);

    for (let i = 1; i <= steps; i++) {
      const x = lo + i * dx;
      const y = f(x);

      if (isFinite(prevY) && isFinite(y)) {
        if (Math.abs(prevY) < 1e-9) {
          roots.push(prevX);
        } else if (prevY * y < 0) {
          let a = prevX, b = x, fa = prevY;
          for (let k = 0; k < 60; k++) {
            const m = (a + b) / 2;
            const fm = f(m);
            if (!isFinite(fm)) break;
            if (fa * fm <= 0) { b = m; } else { a = m; fa = fm; }
          }
          roots.push((a + b) / 2);
        }
      }
      prevX = x;
      prevY = y;
    }
    const rounded = roots.map(r => Math.round(r * 1e6) / 1e6);
    const uniq: number[] = [];
    rounded.sort((a, b) => a - b).forEach(r => {
      if (uniq.length === 0 || Math.abs(r - uniq[uniq.length - 1]) > 1e-4) uniq.push(r);
    });
    return uniq;
};

// ==========================================
// HÀM CHÍNH: XUẤT DỮ LIỆU CỘT CHO BẢNG BIẾN THIÊN
// ==========================================
export const calculateVariationTableData = (latex: string, ce: any) => {
    let rhs = latex;
    if (latex.includes("=")) {
      rhs = latex.substring(latex.indexOf("=") + 1).trim();
    }
    const expr = ce.parse(rhs);

    let rawDiscontinuities: number[] = [];
    const fracRegex = /\\frac\{([^}]+)\}\{([^}]+)\}/g;
    let match;
    while ((match = fracRegex.exec(rhs)) !== null) {
      const denomExpr = ce.parse(match[2]);
      let solveRootsDenom: number[] = [];
      try {
        const solveDenom = ce.expr(["Solve", ["Equal", denomExpr, 0], "x"]).evaluate();
        extractRootsRecursively(solveDenom, solveRootsDenom);
      } catch { }

      const denomFn = (xVal: number) => {
        try { return getNum(denomExpr.subs({ x: xVal })); }
        catch { return NaN; }
      };
      const numericRootsDenom = findRootsNumeric(denomFn);

      let mergedDenom: number[] = [...solveRootsDenom];
      numericRootsDenom.forEach(nr => {
        if (!mergedDenom.some(r => Math.abs(r - nr) < 1e-3)) mergedDenom.push(nr);
      });
      rawDiscontinuities.push(...mergedDenom);
    }

    const deriv = ce.expr(["Derivative", expr, "x"]).evaluate().simplify();
    let solveRoots: number[] = [];
    try {
      const solveRes = ce.expr(["Solve", ["Equal", deriv, 0], "x"]).evaluate();
      extractRootsRecursively(solveRes, solveRoots);
    } catch { }

    const derivFn = (xVal: number) => {
      try { return getNum(deriv.subs({ x: xVal })); }
      catch { return NaN; }
    };
    const numericRoots = findRootsNumeric(derivFn);

    let rawRoots: number[] = [...solveRoots];
    numericRoots.forEach(nr => {
      if (!rawRoots.some(r => Math.abs(r - nr) < 1e-3)) rawRoots.push(nr);
    });

    let allPoints = [...rawRoots, ...rawDiscontinuities].sort((a, b) => a - b);
    let uniquePoints: { val: number, isDis: boolean }[] = [];
    for (let i = 0; i < allPoints.length; i++) {
      if (i > 0 && Math.abs(allPoints[i] - allPoints[i - 1]) < 1e-6) continue;
      const val = allPoints[i];
      const isDis = rawDiscontinuities.some(d => Math.abs(d - val) < 1e-6);
      uniquePoints.push({ val, isDis });
    }

    const signOfDerivAt = (xVal: number) => {
      const res = getNum(deriv.subs({ x: xVal }));
      if (isNaN(res) || !isFinite(res)) return "+"; 
      return res >= 0 ? "+" : "-";
    };

    let signs: string[] = [];
    for (let i = 0; i <= uniquePoints.length; i++) {
      let testPoint = 0;
      if (i === 0) testPoint = uniquePoints.length > 0 ? uniquePoints[0].val - 1 : 0;
      else if (i === uniquePoints.length) testPoint = uniquePoints[i - 1].val + 1;
      else testPoint = (uniquePoints[i - 1].val + uniquePoints[i].val) / 2;
      signs.push(signOfDerivAt(testPoint));
    }

    const getLimit = (xVal: number, direction = 0) => {
      if (xVal === -1000 || xVal === 1000) {
        const testX = xVal === -1000 ? -1e7 : 1e7;
        const res = getNum(expr.subs({ x: testX }));
        if (isNaN(res)) return "?";
        if (!isFinite(res) || Math.abs(res) > 1e6) return res > 0 ? "+\\infty" : "-\\infty";
        return niceNumber(res);
      }
      if (direction === 0) {
        const val = getNum(expr.subs({ x: xVal }));
        if (isNaN(val)) return "?";
        if (!isFinite(val)) return val > 0 ? "+\\infty" : "-\\infty";
        return niceNumber(val);
      }

      const epsBig = 1e-3;
      const epsSmall = 1e-6;
      const valBig = getNum(expr.subs({ x: xVal + direction * epsBig }));
      const valSmall = getNum(expr.subs({ x: xVal + direction * epsSmall }));

      if (!isFinite(valSmall) || isNaN(valSmall)) return valSmall > 0 ? "+\\infty" : "-\\infty";
      if (isFinite(valBig) && Math.abs(valSmall) > Math.abs(valBig) * 20) return valSmall > 0 ? "+\\infty" : "-\\infty";
      if (Math.abs(valSmall) > 1e5) return valSmall > 0 ? "+\\infty" : "-\\infty";

      return niceNumber(valSmall);
    };

    let cols: any[] = [];
    cols.push({ type: 'val', x: "-\\infty", isDis: false, yPrime: "", y: getLimit(-1000) });

    for (let i = 0; i < uniquePoints.length; i++) {
      let pt = uniquePoints[i];
      cols.push({ type: 'interval', sign: signs[i] });

      if (pt.isDis) {
        cols.push({
          type: 'val', x: niceNumber(pt.val), isDis: true, yPrime: "",
          yLeft: getLimit(pt.val, -1), yRight: getLimit(pt.val, 1)
        });
      } else {
        cols.push({
          type: 'val', x: niceNumber(pt.val), isDis: false, yPrime: "0", y: getLimit(pt.val, 0)
        });
      }
    }

    cols.push({ type: 'interval', sign: signs[signs.length - 1] });
    cols.push({ type: 'val', x: "+\\infty", isDis: false, y: getLimit(1000) });

    return cols;
};