// ==========================================
// VẼ SVG PHÉP TÍNH DỌC
// ==========================================
export const buildVerticalMathPreview = (bieuThuc: string) => {
    const laPhepNhan = bieuThuc.includes('\\times') || bieuThuc.includes('\\cdot') || bieuThuc.includes('*') || bieuThuc.includes('x');
    const laPhepChia = bieuThuc.includes('\\frac') || bieuThuc.includes('\\div') || bieuThuc.includes('/');
    const laPhepCong = bieuThuc.includes('+');
    const laPhepTru = bieuThuc.includes('-');

    if (!laPhepNhan && !laPhepChia && !laPhepCong && !laPhepTru) return null;

    const CHAR_W = 16.5; const LINE_H = 32; const MARGIN = 15;
    const FONT_STYLE = `font-family="Arial, monospace" font-size="28px"`;

    let svgString = "";
    let mathWidthEx = 0;
    let mathHeightEx = 0;

    if (laPhepNhan) {
        let chuoiSach = bieuThuc.replace(/\\times/g, '*').replace(/\\cdot/g, '*').replace(/x/g, '*').replace(/\\text\{\*\}/g, '*');
        const mangSo = chuoiSach.split('*');
        if (mangSo.length !== 2) return null;

        let strSo1 = mangSo[0].trim(); let strSo2 = mangSo[1].trim();
        let nSo1 = parseInt(strSo1, 10); let nSo2 = parseInt(strSo2, 10);
        if (isNaN(nSo1) || isNaN(nSo2)) return null;

        let ketQuaCuoi = (nSo1 * nSo2).toString();
        const maxChars = Math.max(strSo1.length, strSo2.length, ketQuaCuoi.length);
        const width = maxChars * CHAR_W + MARGIN * 3;

        let linesCount = 2; let partials = [];
        if (strSo2.length > 1) {
            for (let i = strSo2.length - 1; i >= 0; i--) {
                partials.push((nSo1 * parseInt(strSo2[i], 10)).toString());
            }
            linesCount += partials.length + 1;
        } else { linesCount += 1; }

        const height = linesCount * LINE_H + MARGIN * 2;
        mathWidthEx = width / 14; mathHeightEx = height / 14;

        svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;
        const rightEdge = width - MARGIN;

        svgString += `<text x="${rightEdge}" y="${LINE_H + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo1}</text>`;
        svgString += `<text x="${rightEdge}" y="${LINE_H * 2 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo2}</text>`;
        let offsetDauNhanX = rightEdge - Math.max(strSo1.length, strSo2.length) * CHAR_W - 10;
        svgString += `<text x="${offsetDauNhanX}" y="${LINE_H * 1.5 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">×</text>`;
        svgString += `<line x1="${MARGIN}" y1="${LINE_H * 2 + MARGIN + 8}" x2="${rightEdge}" y2="${LINE_H * 2 + MARGIN + 8}" stroke="black" stroke-width="3"/>`;

        let currentY = LINE_H * 3 + MARGIN + 6;
        if (partials.length === 0) {
            svgString += `<text x="${rightEdge}" y="${currentY}" text-anchor="end" ${FONT_STYLE} fill="black">${ketQuaCuoi}</text>`;
        } else {
            partials.forEach((p, index) => {
                let offsetRight = rightEdge - (index * CHAR_W);
                svgString += `<text x="${offsetRight}" y="${currentY}" text-anchor="end" ${FONT_STYLE} fill="black">${p}</text>`;
                currentY += LINE_H;
            });
            svgString += `<line x1="${MARGIN}" y1="${currentY - LINE_H + 12}" x2="${rightEdge}" y2="${currentY - LINE_H + 12}" stroke="black" stroke-width="3"/>`;
            svgString += `<text x="${rightEdge}" y="${currentY + 8}" text-anchor="end" ${FONT_STYLE} fill="black">${ketQuaCuoi}</text>`;
        }
        svgString += `</svg>`;

    } else if (laPhepCong || laPhepTru) {
        let toanTu = laPhepCong ? '+' : '-';
        const mangSo = bieuThuc.split(toanTu);
        if (mangSo.length !== 2) return null;

        let strSo1 = mangSo[0].trim(); let strSo2 = mangSo[1].trim();
        let nSo1 = parseInt(strSo1, 10); let nSo2 = parseInt(strSo2, 10);
        if (isNaN(nSo1) || isNaN(nSo2)) return null;

        let ketQuaCuoi = laPhepCong ? (nSo1 + nSo2).toString() : (nSo1 - nSo2).toString();
        const maxChars = Math.max(strSo1.length, strSo2.length, ketQuaCuoi.length);
        const width = maxChars * CHAR_W + MARGIN * 3;
        const height = 3 * LINE_H + MARGIN * 2; 

        mathWidthEx = width / 14; mathHeightEx = height / 14;

        svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;
        const rightEdge = width - MARGIN;

        svgString += `<text x="${rightEdge}" y="${LINE_H + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo1}</text>`;
        svgString += `<text x="${rightEdge}" y="${LINE_H * 2 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo2}</text>`;

        let offsetDauX = rightEdge - Math.max(strSo1.length, strSo2.length) * CHAR_W - 10;
        svgString += `<text x="${offsetDauX}" y="${LINE_H * 1.5 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${toanTu}</text>`;
        svgString += `<line x1="${MARGIN}" y1="${LINE_H * 2 + MARGIN + 8}" x2="${rightEdge}" y2="${LINE_H * 2 + MARGIN + 8}" stroke="black" stroke-width="3"/>`;
        let currentY = LINE_H * 3 + MARGIN + 6;
        svgString += `<text x="${rightEdge}" y="${currentY + 8}" text-anchor="end" ${FONT_STYLE} fill="black">${ketQuaCuoi}</text>`;
        svgString += `</svg>`;

    } else {
        let strBiChia = ""; let strChia = "";
        if (bieuThuc.includes('\\frac')) {
            const match = bieuThuc.match(/\\frac\{([0-9]+)\}\{([0-9]+)\}/);
            if (match) { strBiChia = match[1]; strChia = match[2]; }
        } else {
            let chuoiSach = bieuThuc.replace(/\\div/g, '/').replace(/\\text\{\/\}/g, '/');
            const mangSo = chuoiSach.split('/');
            if (mangSo.length === 2) { strBiChia = mangSo[0].trim(); strChia = mangSo[1].trim(); }
        }

        const nBiChia = parseInt(strBiChia, 10);
        const nChia = parseInt(strChia, 10);
        if (isNaN(nBiChia) || isNaN(nChia) || nChia === 0) return null;

        let thuong = "";
        let leftRows = [];
        leftRows.push({ text: strBiChia, spaces: 0 });

        let soDuStr = ""; let chiSo = 0; let laLuotDauTien = true;
        while (chiSo < strBiChia.length) {
            if (laLuotDauTien) {
                let soTamStr = strBiChia[chiSo];
                while (parseInt(soTamStr, 10) < nChia && chiSo + 1 < strBiChia.length) {
                    chiSo++; soTamStr += strBiChia[chiSo];
                }
                let nTam = parseInt(soTamStr, 10);
                thuong += Math.floor(nTam / nChia).toString();
                soDuStr = (nTam % nChia).toString();
                laLuotDauTien = false; chiSo++;
            } else {
                let chuSoHa = strBiChia[chiSo];
                let dongHaSo = soDuStr + chuSoHa;
                let viTriHa = chiSo + 1;
                leftRows.push({ text: dongHaSo, spaces: Math.max(0, viTriHa - dongHaSo.length) });

                let nTam = parseInt(dongHaSo, 10);
                thuong += Math.floor(nTam / nChia).toString();
                soDuStr = (nTam % nChia).toString();
                chiSo++;
            }
        }
        let viTriCuoi = strBiChia.length;
        leftRows.push({ text: soDuStr, spaces: Math.max(0, viTriCuoi - soDuStr.length) });

        const leftMaxChars = strBiChia.length;
        const rightMaxChars = Math.max(strChia.length, thuong.length);
        const width = (leftMaxChars + rightMaxChars + 2) * CHAR_W + MARGIN * 3;
        const height = Math.max(leftRows.length, 2) * LINE_H + MARGIN * 2;

        mathWidthEx = width / 14; mathHeightEx = height / 14;

        svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;
        const vLineX = MARGIN + leftMaxChars * CHAR_W + (CHAR_W / 2);
        svgString += `<line x1="${vLineX}" y1="${MARGIN / 2}" x2="${vLineX}" y2="${height - MARGIN}" stroke="black" stroke-width="3"/>`;

        leftRows.forEach((row, idx) => {
            let xPos = MARGIN + row.spaces * CHAR_W;
            let yPos = (idx + 1) * LINE_H + MARGIN / 2;
            svgString += `<text x="${xPos}" y="${yPos}" ${FONT_STYLE} fill="black">${row.text}</text>`;
        });

        const rightX = vLineX + CHAR_W;
        svgString += `<text x="${rightX}" y="${LINE_H + MARGIN / 2}" ${FONT_STYLE} fill="black">${strChia}</text>`;
        svgString += `<line x1="${vLineX}" y1="${LINE_H + MARGIN / 2 + 8}" x2="${width - MARGIN}" y2="${LINE_H + MARGIN / 2 + 8}" stroke="black" stroke-width="3"/>`;
        svgString += `<text x="${rightX}" y="${LINE_H * 2 + MARGIN / 2 + 4}" ${FONT_STYLE} fill="#010101">${thuong || "0"}</text>`;
        svgString += `</svg>`;
    }

    return { svgString, mathWidthEx, mathHeightEx, bieuThuc };
};


// ==========================================
// VẼ SVG BẢNG BIẾN THIÊN (CÓ SCALE ĐÚNG TỶ LỆ)
// ==========================================
export const buildVariationTableSvg = (bbtData: any) => {
    const { cols: rawCols, labels: rawLabels } = bbtData;
    const cols = JSON.parse(JSON.stringify(rawCols));

    let valCols = cols.filter((c: any) => c.type === 'val');
    let signs = cols.filter((c: any) => c.type === 'interval').map((c: any) => c.sign);

    for (let i = 0; i < valCols.length; i++) {
        let c = valCols[i];
        if (i === 0) {
            c.pos = signs[0] === '+' || signs[0] === '-' ? (signs[0] === '+' ? 'BOTTOM' : 'TOP') : 'MIDDLE';
        } else if (i === valCols.length - 1) {
            c.pos = signs[signs.length - 1] === '+' || signs[signs.length - 1] === '-' ? (signs[signs.length - 1] === '+' ? 'TOP' : 'BOTTOM') : 'MIDDLE';
        } else {
            let inSign = signs[i - 1];
            let outSign = signs[i];

            if (c.isDis) {
                c.posLeft = inSign === '+' ? 'TOP' : 'BOTTOM';
                c.posRight = outSign === '+' ? 'BOTTOM' : 'TOP';
            } else {
                if (inSign === '+' && outSign === '-') c.pos = 'TOP';
                else if (inSign === '-' && outSign === '+') c.pos = 'BOTTOM';
                else c.pos = 'MIDDLE';
            }
        }
    }

    const X_OFFSET = 80;  
    const COL_W = 120; 
    const RIGHT_PADDING = 10; 
    const W = X_OFFSET + cols.length * COL_W + RIGHT_PADDING; 
    const H = 290; 

    let currentX = X_OFFSET;
    cols.forEach((c: any) => { c.cx = currentX + COL_W / 2; currentX += COL_W; });

    let svgString = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="100%" height="auto" preserveAspectRatio="xMidYMid meet" style="background-color: white; max-width: ${W}px; max-height: ${H}px; display: block; margin: 0 auto;">`;
    svgString += `<defs><marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="black" /></marker></defs>`;
    svgString += `<rect x="0" y="0" width="${W}" height="${H}" fill="none" stroke="black" stroke-width="1.5"/>`;
    
    svgString += `<line x1="0" y1="60" x2="${W}" y2="60" stroke="black" stroke-width="1.5"/>`;
    svgString += `<line x1="0" y1="105" x2="${W}" y2="105" stroke="black" stroke-width="1.5"/>`;
    svgString += `<line x1="${X_OFFSET}" y1="0" x2="${X_OFFSET}" y2="${H}" stroke="black" stroke-width="1.5"/>`;

    const FONT = `font-family="Times New Roman, serif" font-size="28px" text-anchor="middle" dominant-baseline="middle"`;

    const renderCell = (text: string, cx: number, cy: number, isBold = false, isRed = false, isItalic = false) => {
        let rawText = (text || "").toString().trim();
        if (!rawText) return "";
        
        const isLatex = rawText.includes('\\') || rawText.includes('^') || rawText.includes('_');
        
        if (isLatex) {
            try {
                const mathJax = (window as any).MathJax;
                if (mathJax && mathJax.tex2svg) {
                    const node = mathJax.tex2svg(rawText);
                    const svg = node.querySelector("svg");
                    if (svg) {
                        const wStr = svg.getAttribute("width") || "0";
                        const hStr = svg.getAttribute("height") || "0";
                        const viewBox = svg.getAttribute("viewBox") || "";
                        
                        const EX_TO_PX = 10; 
                        let w = parseFloat(wStr) * EX_TO_PX;
                        let h = parseFloat(hStr) * EX_TO_PX;
                        
                        const x = cx - w / 2;
                        const y = cy - h / 2;
                        
                        let color = isRed ? "red" : (isBold ? "#000" : "black");
                        let inner = svg.innerHTML.replace(/currentColor/gi, color);
                        
                        return `<svg x="${x}" y="${y}" width="${w}" height="${h}" viewBox="${viewBox}">${inner}</svg>`;
                    }
                }
            } catch (e) { console.warn("Lỗi render LaTeX BBT:", e); }
        }

        let cleanText = rawText.replace(/-\\infty/g, "-\u221E").replace(/\+\\infty/g, "+\u221E").replace(/\\infty/g, "\u221E");
        const weight = isBold ? 'font-weight="bold"' : '';
        const fill = isRed ? 'fill="red"' : 'fill="black"';
        const italic = isItalic ? 'font-style="italic"' : '';
        return `<text x="${cx}" y="${cy}" ${FONT} ${weight} ${fill} ${italic}>${cleanText}</text>`;
    };

    const labels = rawLabels || { x: "x", yPrime: "y'", y: "y" };
    svgString += renderCell(labels.x, X_OFFSET/2, 30, false, false, true);
    svgString += renderCell(labels.yPrime, X_OFFSET/2, 82.5, false, false, true);
    svgString += renderCell(labels.y, X_OFFSET/2, 197.5, false, false, true);

    const Y_TOP = 140;
    const Y_MID = 197.5;
    const Y_BOT = 255;

    cols.forEach((c: any, i: number) => {
        let shiftX = 0;
        if (i === 0) shiftX = -20;
        else if (i === cols.length - 1) shiftX = 10;
        
        let currentCx = c.cx + shiftX;

        if (c.type === 'interval') {
            const prevShift = (i - 1 === 0) ? -20 : 0;
            const nextShift = (i + 1 === cols.length - 1) ? 10 : 0;
            currentCx = (cols[i - 1].cx + prevShift + cols[i + 1].cx + nextShift) / 2;
        }

        if (c.type === 'val') {
            svgString += renderCell(c.x, currentCx, 30);
            
            if (c.isDis) {
                svgString += `<line x1="${c.cx - 3}" y1="60" x2="${c.cx - 3}" y2="${H}" stroke="black" stroke-width="1.5"/>`;
                svgString += `<line x1="${c.cx + 3}" y1="60" x2="${c.cx + 3}" y2="${H}" stroke="black" stroke-width="1.5"/>`;
                
                svgString += renderCell(c.yLeft, c.cx - 35, c.posLeft === 'TOP' ? Y_TOP : Y_BOT, false);
                svgString += renderCell(c.yRight, c.cx + 35, c.posRight === 'TOP' ? Y_TOP : Y_BOT, false);
            } else {
                if (c.yPrime) svgString += renderCell(c.yPrime, currentCx, 82.5);
                svgString += renderCell(c.y, currentCx, c.pos === 'TOP' ? Y_TOP : (c.pos === 'BOTTOM' ? Y_BOT : Y_MID), false);
            }
        } else {
            svgString += renderCell(c.sign, currentCx, 82.5, false, false); 
            
            if (c.sign === '+' || c.sign === '-') {
                const prevVal = cols[i - 1];
                const nextVal = cols[i + 1];
                
                let prevShift = (i - 1 === 0) ? -20 : 0;
                let nextShift = (i + 1 === cols.length - 1) ? 10 : 0;

                let y1 = prevVal.isDis ? (prevVal.posRight === 'TOP' ? Y_TOP : Y_BOT) : (prevVal.pos === 'TOP' ? Y_TOP : (prevVal.pos === 'BOTTOM' ? Y_BOT : Y_MID));
                let y2 = nextVal.isDis ? (nextVal.posLeft === 'TOP' ? Y_TOP : Y_BOT) : (nextVal.pos === 'TOP' ? Y_TOP : (nextVal.pos === 'BOTTOM' ? Y_BOT : Y_MID));
                
                const PADDING_X_START = prevVal.isDis ? 65 : 45;
                const PADDING_X_END = nextVal.isDis ? 65 : 55;

                let x1 = prevVal.cx + prevShift + PADDING_X_START;
                let x2 = nextVal.cx + nextShift - PADDING_X_END;
                
                svgString += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="black" stroke-width="1.5" marker-end="url(#arrow)" />`;
            }
        }
    });

    svgString += `</svg>`;
    return { svgString, mathWidthEx: W/14, mathHeightEx: H/14 };
};