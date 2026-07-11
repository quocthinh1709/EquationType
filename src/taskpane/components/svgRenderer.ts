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


// ==========================================
// VẼ SVG ĐỒ THỊ HÀM SỐ (GRAPH) - CÓ MÀU VÀ TÊN ĐIỂM
// ==========================================
export const buildGraphSvg = (
    evaluator: (x: number) => number, 
    config: { xMin: number, xMax: number, yMin: number, yMax: number },
    annotations: { name: string, x: number, y: number }[], // Thêm trường name
    lineColor: string = "#d13438" // Thêm tham số màu (Mặc định là đỏ)
) => {
    const { xMin, xMax, yMin, yMax } = config;
    const W = 600; 
    const H = 600;
    const P = 50; // Padding lề

    const mapX = (x: number) => P + ((x - xMin) / (xMax - xMin)) * (W - 2 * P);
    const mapY = (y: number) => H - P - ((y - yMin) / (yMax - yMin)) * (H - 2 * P);

    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="100%" height="auto" preserveAspectRatio="xMidYMid meet" style="background-color: white; max-width: ${W}px; max-height: ${H}px; display: block; margin: 0 auto; font-family: 'Times New Roman', serif;">`;
    
    //svg += `<defs><marker id="axis-arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="black" /></marker></defs>`;
    svg += `<defs>
                <marker id="axis-arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
                    <path d="M 0 0 L 10 5 L 0 10 z" fill="black" />
                </marker>
                <clipPath id="graphClip">
                    <rect x="${P}" y="${P}" width="${W - 2 * P}" height="${H - 2 * P}" />
                </clipPath>
            </defs>`;

    const axisColor = "#000000";
    const gridColor = "#bebebd";
    
    for (let i = Math.ceil(xMin); i <= Math.floor(xMax); i++) {
        const cx = mapX(i);
        if (i !== 0) {
            svg += `<line x1="${cx}" y1="${P}" x2="${cx}" y2="${H-P}" stroke="${gridColor}" stroke-width="1.2" stroke-dasharray="4 4" />`;
            if (xMax - xMin <= 20 || i % 2 === 0) {
                svg += `<text x="${cx}" y="${mapY(0) + 18}" font-size="22px" text-anchor="middle" fill="black">${i}</text>`;
            }
        }
    }

    for (let i = Math.ceil(yMin); i <= Math.floor(yMax); i++) {
        const cy = mapY(i);
        if (i !== 0) {
            svg += `<line x1="${P}" y1="${cy}" x2="${W-P}" y2="${cy}" stroke="${gridColor}" stroke-width="1.2" stroke-dasharray="4 4" />`;
            if (yMax - yMin <= 20 || i % 2 === 0) {
                svg += `<text x="${mapX(0) - 10}" y="${cy + 5}" font-size="22px" text-anchor="end" fill="black">${i}</text>`;
            }
        }
    }

    const originY = Math.max(P, Math.min(H - P, mapY(0)));
    svg += `<line x1="${P/2}" y1="${originY}" x2="${W - P/2}" y2="${originY}" stroke="${axisColor}" stroke-width="1.5" marker-end="url(#axis-arrow)" />`;
    svg += `<text x="${W - P/2 + 10}" y="${originY + 5}" font-size="22px" font-style="italic">x</text>`;

    const originX = Math.max(P, Math.min(W - P, mapX(0)));
    svg += `<line x1="${originX}" y1="${H - P/2}" x2="${originX}" y2="${P/2}" stroke="${axisColor}" stroke-width="1.5" marker-end="url(#axis-arrow)" />`;
    svg += `<text x="${originX + 10}" y="${P/2 - 10}" font-size="22px" font-style="italic">y</text>`;
    
    svg += `<text x="${originX - 22}" y="${originY + 22}" font-size="22px" font-style="italic">O</text>`;

    const steps = 400;
    const dx = (xMax - xMin) / steps;
    let pathD = "";
    let prevY = NaN;

    for (let i = 0; i <= steps; i++) {
        const x = xMin + i * dx;
        const y = evaluator(x);

        if (!isNaN(y) && isFinite(y)) {
            const cx = mapX(x);
            const cy = mapY(y);

            const jump = Math.abs(y - prevY);
            if (i === 0 || isNaN(prevY) || jump > (yMax - yMin) / 2) {
                pathD += ` M ${cx} ${cy}`; 
            } else {
                pathD += ` L ${cx} ${cy}`; 
            }
            prevY = y;
        } else {
            prevY = NaN;
        }
    }
    
    // ĐÃ ĐỔI: Dùng biến lineColor thay vì màu đỏ cố định
    svg += `<path d="${pathD}" fill="none" stroke="${lineColor}" stroke-width="2.5" stroke-linejoin="round" clip-path="url(#graphClip)" />`;

    annotations.forEach(pt => {
        const cx = mapX(pt.x);
        const cy = mapY(pt.y);
        
        svg += `<line x1="${cx}" y1="${originY}" x2="${cx}" y2="${cy}" stroke="black" stroke-dasharray="4 4" stroke-width="1.5" />`;
        svg += `<line x1="${originX}" y1="${cy}" x2="${cx}" y2="${cy}" stroke="black" stroke-dasharray="4 4" stroke-width="1.5" />`;
        
        // Chấm tròn điểm tọa độ cũng dùng màu cùng với đường đồ thị cho đồng bộ
        svg += `<circle cx="${cx}" cy="${cy}" r="4" fill="${lineColor}" />`;
        
        // ĐÃ ĐỔI: Thêm tên điểm (VD: A) vào nhãn
        const prefix = pt.name ? `${pt.name}` : "";
        const label = `${prefix}(${Math.round(pt.x*100)/100}, ${Math.round(pt.y*100)/100})`;
        
        svg += `<text x="${cx + 8}" y="${cy - 8}" font-size="22px" font-weight="bold" fill="black">${label}</text>`;
    });

    svg += `</svg>`;
    return { svgString: svg, mathWidthEx: W/14, mathHeightEx: H/14 };
};