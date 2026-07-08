import * as React from "react";
import { useEffect, useRef, useState, useCallback, useMemo } from "react";
import { DefaultButton, IconButton, Stack, Label, ChoiceGroup, IChoiceGroupOption, Toggle, MessageBar, MessageBarType } from "@fluentui/react";
import { Dialog, DialogType, DialogFooter, PrimaryButton, TextField, Panel, PanelType, SearchBox, ThemeProvider, createTheme } from "@fluentui/react";
import { initializeIcons } from '@fluentui/react/lib/Icons';

import { MathfieldElement } from "mathlive";
import { ComputeEngine } from "@cortex-js/compute-engine";
import "mathlive";
import MathPalette from "./MathPalette";
import { customInlineShortcuts, customMacros } from "./mathShortcuts";
import { unsupportedMathMlCommands } from "./unsupportedCommands";

initializeIcons();

// Khởi tạo engine
const ce = new ComputeEngine();

const officeGreenTheme = createTheme({
  palette: {
    themePrimary: '#107c41',
    themeLighterAlt: '#f2fafa',
    themeLighter: '#cff0db',
    themeLight: '#aae5be',
    themeTertiary: '#65cb8b',
    themeSecondary: '#299155',
    themeDarkAlt: '#0e703a',
    themeDark: '#0c5e31',
    themeDarker: '#094624',
  }
});

MathfieldElement.fontsDirectory = "./fonts";

export default function App() {
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const lastEqIdRef = useRef<string>("");
  const [currentLatex, setCurrentLatex] = useState("");
  const [insertMode, setInsertMode] = useState<string>("mathml");


  // --- BẮT ĐẦU: STATE VÀ LOGIC CHO COMPUTE ENGINE ---
  const [computeResult, setComputeResult] = useState<string>("");

  const handleCompute = (action: "evaluate" | "simplify" | "expand" | "expandAll"| "factor" | "solve") => {
    if (!currentLatex) return;

    try {
      const expr = ce.parse(currentLatex);
      let resultLatex = "";

      switch (action) {
        case "evaluate": // Tính toán (vd: 1/2 + 1/3 = 5/6)
          resultLatex = expr.evaluate().latex;
          break;
        case "simplify": // Rút gọn (vd: 2x + x = 3x)
          resultLatex = expr.simplify().latex;
          break;
        case "expand": // Khai triển (vd: (x+1)^2 = x^2 + 2x + 1)
          resultLatex = ce.expr(["Expand", expr]).evaluate().latex;
          break;
        case "expandAll": // TÍNH NĂNG MỚI: Khai triển triệt để
          resultLatex = ce.expr(["ExpandAll", expr]).evaluate().latex;
          break;
        case "factor": // Phân tích nhân tử (vd: x^2 - 1 = (x-1)(x+1))
          resultLatex = ce.expr(["Factor", expr]).evaluate().latex;
          break;
        case "solve": 
          // 1. Tự động tìm các ẩn số có trong biểu thức (VD: x, y, z...)
          let vars: string[] = [];
          if (expr.unknowns && expr.unknowns.length > 0) {
            vars = [...expr.unknowns]; // Lấy danh sách ẩn tự động từ CortexJS
          } else {
            // Backup tự dò nếu hệ thống nhận diện sót
            if (currentLatex.includes("x")) vars.push("x");
            if (currentLatex.includes("y")) vars.push("y");
            if (currentLatex.includes("z")) vars.push("z");
            if (vars.length === 0) vars = ["x"]; // Mặc định là x
          }

          // 2. Phân loại: Hệ phương trình vs Phương trình 1 ẩn
          if (vars.length > 1) {
             // ---> GIẢI HỆ PHƯƠNG TRÌNH NHIỀU ẨN <---
             // Với hệ, hàm expr.solve() tích hợp sẵn sẽ trả về đối tượng { x: ..., y: ... }
             if (typeof expr.solve === "function") {
                const solution = expr.solve(vars) as Record<string, any>;
                
                if (solution && typeof solution === "object" && !Array.isArray(solution) && Object.keys(solution).length > 0) {
                   // Dựng lại mã LaTeX hệ phương trình kết quả
                   let casesContent = vars.map(v => {
                      return solution[v] ? `${v} = ${solution[v].latex}` : "";
                   }).filter(Boolean).join(" \\\\ "); // Nối các dòng bằng xuống dòng (\\)
                   
                   resultLatex = `\\displaystyle \\begin{cases} ${casesContent} \\end{cases}`;
                } else {
                   resultLatex = "\\text{Hệ vô nghiệm hoặc số liệu không hợp lệ}";
                }
             } else {
                resultLatex = "\\text{Phiên bản thư viện chưa hỗ trợ hàm này}";
             }
          } else {
             // ---> GIẢI PHƯƠNG TRÌNH 1 ẨN (Giữ nguyên code cũ đang chạy tốt) <---
             const v = vars[0];
             const solveResult = ce.expr(["Solve", expr, v]).evaluate();
             
             if (solveResult && solveResult.latex.includes("Solve")) {
               resultLatex = "\\text{Chưa hỗ trợ giải phương trình này hoặc vô nghiệm}";
             } else {
               resultLatex = `${v} \\in \\{` + solveResult.latex + "\\}";
               // Dọn dẹp dấu ngoặc nhọn bị lặp
               resultLatex = resultLatex.replace(/\\\{\\\{/g, "\\{").replace(/\\\}\\\}/g, "\\}");
             }
          }
          break;
      }

      setComputeResult(resultLatex);
    } catch (error) {
      console.error("Lỗi Compute Engine:", error);
      setComputeResult("Lỗi: Không thể tính toán biểu thức này.");
    }
  };

  const handleInsertComputeResult = () => {
    if (!computeResult || !mathfieldRef.current) return;

    // Đẩy kết quả tính toán ngược lại vào ô MathLive chính
    mathfieldRef.current.value = computeResult;
    setCurrentLatex(computeResult);

    // Dùng setTimeout 100ms để đảm bảo React đã cập nhật biến currentLatex trước khi chèn vào Word
    setTimeout(() => {
      handleInsertEquation();
    }, 100);
  };
  // --- KẾT THÚC: LOGIC COMPUTE ENGINE ---

  const [userDict, setUserDict] = useState<{ inline: Record<string, string>, macro: Record<string, string> }>(() => {
    try {
      const saved = localStorage.getItem("math_user_dict");
      return saved ? JSON.parse(saved) : { inline: {}, macro: {} };
    } catch (e) {
      return { inline: {}, macro: {} };
    }
  });

  const [isAddDictOpen, setIsAddDictOpen] = useState(false);
  const [dictKey, setDictKey] = useState("");
  const [dictType, setDictType] = useState<"inline" | "macro">("inline");
  const fileInputRef = useRef<HTMLInputElement>(null);

  interface EditItem { id: string; type: "inline" | "macro"; key: string; value: string; }
  const [isEditDictOpen, setIsEditDictOpen] = useState(false);
  const [editItems, setEditItems] = useState<EditItem[]>([]);
  const [searchTerm, setSearchTerm] = useState("");

  const isUnsupported = unsupportedMathMlCommands.some(cmd => currentLatex.includes(cmd));

  useEffect(() => {
    if (isUnsupported && insertMode === "mathml") {
      setInsertMode("svg");
    }
  }, [isUnsupported, insertMode]);

  const upgradeMacros = (macroDict: Record<string, any>) => {
    const upgraded: Record<string, any> = {};
    Object.keys(macroDict).forEach(key => {
      const val = macroDict[key];
      if (typeof val === "string") {
        upgraded[key] = {
          def: val,
          captureSelection: false,
          expand: true
        };
      } else {
        upgraded[key] = val;
      }
    });
    return upgraded;
  };

  // ==========================================
  // HÀM 1: ĐỌC NGƯỢC TỪ WORD VÀO ADD-IN VÀ DỌN RÁC 
  // (ĐÃ NÂNG CẤP: "SMART ALT TEXT" - TRÍCH XUẤT ID THÔNG MINH)
  // ==========================================
  const handleReadEquation = useCallback(async () => {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();

        const cc = range.parentContentControlOrNullObject;
        cc.load("title, tag");
        const ccs = range.contentControls;
        ccs.load("items/title, items/tag");
        
        // Tải Thuộc tính ẢNH để dự phòng khi Copy/Paste làm mất Khung
        const pics = range.inlinePictures;
        pics.load("items/altTextDescription");

        await context.sync();

        let newFullTag = "";
        let newTargetCC: Word.ContentControl | null = null;

        // ƯU TIÊN 1: Đọc từ Khung (Nếu khung còn nguyên vẹn)
        if (ccs.items.length > 0 && ccs.items[0].title === "LSTS_Math") {
          newFullTag = ccs.items[0].tag;
          newTargetCC = ccs.items[0];
        } else if (!cc.isNullObject && cc.title === "LSTS_Math") {
          newFullTag = cc.tag;
          newTargetCC = cc;
        } 
        // ƯU TIÊN 2: Trích xuất ID từ Smart Alt Text của Ảnh
        else if (pics.items.length > 0 && pics.items[0].altTextDescription && pics.items[0].altTextDescription.includes("#[ID:")) {
          const desc = pics.items[0].altTextDescription;
          // Dùng RegEx tìm đoạn nằm giữa #[ID: và ]
          const match = desc.match(/#\[ID:(.*?)\]/);
          if (match && match[1]) {
            newFullTag = match[1];
          }
        }

        const newEqId = newFullTag ? newFullTag.split("|")[0] : "";
        const newEqType = newFullTag ? (newFullTag.split("|")[1] || "SVG") : "";

        // --- BƯỚC 1: DỌN DẸP VỎ RÁC CỦA CÔNG THỨC TRƯỚC ĐÓ ---
        if (lastEqIdRef.current && lastEqIdRef.current !== newFullTag) {
          const oldControls = context.document.contentControls.getByTag(lastEqIdRef.current);
          oldControls.load("items");
          await context.sync();

          if (oldControls.items.length > 0) {
            const oldCC = oldControls.items[0];
            const oldRange = oldCC.getRange();
            const oldPics = oldRange.inlinePictures;
            oldPics.load("items");
            oldRange.load("text");
            await context.sync();

            const oldEqType = lastEqIdRef.current.split("|")[1] || "SVG";
            const rawText = oldRange.text;
            const cleanText = rawText.replace(/[\s\u200B-\u200D\uFEFF\uFFFC]/g, '').toLowerCase();
            const isPlaceholder = cleanText.includes("clickortap") || cleanText.includes("vănbản");

            let isOldEmptied = false;
            if (oldEqType === "SVG") {
              if (oldPics.items.length === 0) isOldEmptied = true;
            } else { 
              if (rawText === "" || isPlaceholder) isOldEmptied = true;
            }

            if (isOldEmptied) {
              oldCC.delete(false);
              await context.sync();
            }
          }
        }

        // --- BƯỚC 2: TẢI CÔNG THỨC LÊN MATHLIVE ---
        if (newFullTag) {
          lastEqIdRef.current = newFullTag; 
          let isEmptied = false;
          
          if (newTargetCC) {
             const ccRange = newTargetCC.getRange();
             const ccPics = ccRange.inlinePictures;
             ccPics.load("items");
             ccRange.load("text");
             await context.sync();

             const rawText = ccRange.text;
             const cleanText = rawText.replace(/[\s\u200B-\u200D\uFEFF\uFFFC]/g, '').toLowerCase();
             const isPlaceholder = cleanText.includes("clickortap") || cleanText.includes("vănbản");
             
             if (newEqType === "SVG") {
               isEmptied = ccPics.items.length === 0;
             } else {
               isEmptied = rawText === "" || isPlaceholder;
             }
          }

          if (isEmptied && newTargetCC) {
            // XÓA KHUNG RỖNG
            if (mathfieldRef.current) mathfieldRef.current.value = "";
            setCurrentLatex("");
            lastEqIdRef.current = "";
            newTargetCC.delete(false); 
            await context.sync();
          } else {
            // TẢI DỮ LIỆU TỪ DATABASE
            const latexCode = Office.context.document.settings.get(newEqId);
            if (latexCode && mathfieldRef.current) {
              if (mathfieldRef.current.value !== latexCode) {
                mathfieldRef.current.value = latexCode;
                setCurrentLatex(latexCode);
              }
            }
          }
        } else {
          lastEqIdRef.current = "";
        }
      });
    } catch (error) {
      console.error("Lỗi đọc công thức:", error);
    }
  }, []);
/*
  const handleReadEquation = useCallback(async () => {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const cc = range.parentContentControlOrNullObject;
        cc.load("title, tag");
        const ccs = range.contentControls;
        ccs.load("items/title, items/tag");
        await context.sync();

        let newFullTag = "";
        let newTargetCC: Word.ContentControl | null = null;

        if (ccs.items.length > 0 && ccs.items[0].title === "LSTS_Math") {
          newFullTag = ccs.items[0].tag;
          newTargetCC = ccs.items[0];
        } else if (!cc.isNullObject && cc.title === "LSTS_Math") {
          newFullTag = cc.tag;
          newTargetCC = cc;
        }

        const newEqId = newFullTag ? newFullTag.split("|")[0] : "";
        const newEqType = newFullTag ? (newFullTag.split("|")[1] || "SVG") : "";

        if (lastEqIdRef.current && lastEqIdRef.current !== newFullTag) {
          const oldEqId = lastEqIdRef.current.split("|")[0];
          const oldEqType = lastEqIdRef.current.split("|")[1] || "SVG";
          const oldControls = context.document.contentControls.getByTag(lastEqIdRef.current);
          oldControls.load("items");
          await context.sync();

          let isOldDeleted = false;
          if (oldControls.items.length === 0) {
            isOldDeleted = true;
          } else {
            const oldCC = oldControls.items[0];
            const oldRange = oldCC.getRange();
            const oldPics = oldRange.inlinePictures;
            oldPics.load("items");
            oldRange.load("text");
            await context.sync();

            const rawText = oldRange.text;
            if (oldEqType === "SVG") {
              if (oldPics.items.length === 0) {
                isOldDeleted = true;
              }
            } else {
              if (rawText === "") {
                isOldDeleted = true;
              }
            }

            if (isOldDeleted) {
              oldCC.delete(false);
              await context.sync();
            }
          }

          if (isOldDeleted) {
            //Office.context.document.settings.remove(oldEqId);
            //Office.context.document.settings.saveAsync();
            if (!newFullTag) {
              if (mathfieldRef.current) mathfieldRef.current.value = "";
              setCurrentLatex("");
            }
          }
        }

        if (newFullTag && newTargetCC) {
          lastEqIdRef.current = newFullTag;
          const ccRange = newTargetCC.getRange();
          const pics = ccRange.inlinePictures;
          pics.load("items");
          ccRange.load("text");
          await context.sync();

          const rawText = ccRange.text;
          const cleanText = rawText.replace(/[\s\u200B-\u200D\uFEFF\uFFFC]/g, '').toLowerCase();
          const isPlaceholder = cleanText.includes("clickortap") || cleanText.includes("vănbản");

          let isEmptied = false;
          if (newEqType === "SVG") {
            isEmptied = pics.items.length === 0 && (cleanText === "" || isPlaceholder);
          } else {
            isEmptied = isPlaceholder || rawText === "";
          }

          if (isEmptied) {
            if (mathfieldRef.current) mathfieldRef.current.value = "";
            setCurrentLatex("");
            lastEqIdRef.current = "";
            newTargetCC.delete(false);
            //Office.context.document.settings.remove(newEqId);
            //Office.context.document.settings.saveAsync();
            await context.sync();
          } else {
            const latexCode = Office.context.document.settings.get(newEqId);
            if (latexCode && mathfieldRef.current) {
              if (mathfieldRef.current.value !== latexCode) {
                mathfieldRef.current.value = latexCode;
                setCurrentLatex(latexCode);
              }
            }
          }
        } else {
          lastEqIdRef.current = "";
        }
      });
    } catch (error) {
      // Bỏ qua lỗi ngầm
    }
  }, []);
*/
  const handleInsertEquation = async () => {
    if (!mathfieldRef.current) return;

    const latexOriginal = mathfieldRef.current.value;
    const latexExpanded = mathfieldRef.current.getValue("latex-expanded");

    let insertContent = "";
    let isInsertingImage = false;
    let mathWidthEx = 0;
    let mathHeightEx = 0;
    let verticalAlignEx = 0;

    if (insertMode === "mathml") {
      const hiddenDiv = document.createElement("div");
      hiddenDiv.style.display = "none";
      document.body.appendChild(hiddenDiv);
      const tempMf = new MathfieldElement();
      hiddenDiv.appendChild(tempMf);
      tempMf.value = latexExpanded;
      const mathML = tempMf.getValue("math-ml" as any);
      document.body.removeChild(hiddenDiv);
      insertContent = mathML.includes("<math") ? mathML : `<math xmlns="http://www.w3.org/1998/Math/MathML">${mathML}</math>`;
    } else {
      try {
        const mathJaxObj = (window as any).MathJax;
        if (mathJaxObj && mathJaxObj.tex2svg) {
          const svgNode = mathJaxObj.tex2svg(latexExpanded);
          const svgElement = svgNode.querySelector("svg");

          if (svgElement) {
            const FONT_SIZE_PT = 12;
            const EX_TO_PT = FONT_SIZE_PT * 0.5;

            mathWidthEx = parseFloat(svgElement.getAttribute("width") || "0");
            mathHeightEx = parseFloat(svgElement.getAttribute("height") || "0");
            const widthPt = mathWidthEx * EX_TO_PT;
            const heightPt = mathHeightEx * EX_TO_PT;

            const styleStr = svgElement.getAttribute("style") || "";
            const vaMatch = styleStr.match(/vertical-align:\s*([-0-9.]+)ex/);
            if (vaMatch && vaMatch[1]) {
              verticalAlignEx = parseFloat(vaMatch[1]);
            }

            svgElement.setAttribute("width", widthPt + "pt");
            svgElement.setAttribute("height", heightPt + "pt");
            svgElement.setAttribute("xmlns", "http://www.w3.org/2000/svg");
            svgElement.setAttribute("shape-rendering", "geometricPrecision");

            let svgData = new XMLSerializer().serializeToString(svgElement);
            svgData = svgData.replace(/currentColor/gi, "#000000");
            const utf8Encoded = encodeURIComponent(svgData).replace(/%([0-9A-F]{2})/g, (_match, p1) => {
              return String.fromCharCode(parseInt(p1, 16));
            });

            insertContent = window.btoa(utf8Encoded);
            isInsertingImage = true;
          }
        }
      } catch (error) {
        console.error("Lỗi vẽ PNG:", error);
      }
    }

    try {
      const eqId = "EQ" + new Date().getTime() + "_" + Math.random().toString(36).substring(2, 6);
      Office.context.document.settings.set(eqId, latexExpanded);
      await Office.context.document.settings.saveAsync();

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("font/size");
        await context.sync();

        const currentFontSize = range.font.size || 12;

        if (isInsertingImage) {
          if (!insertContent) return;
          const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
          const exToPt = currentFontSize * 0.44;

          if (mathWidthEx > 0 && mathHeightEx > 0) {
            picture.width = mathWidthEx * exToPt;
            picture.height = mathHeightEx * exToPt;
          }
          picture.lockAspectRatio = true;


          // ---> BỔ SUNG: SMART ALT TEXT <---
          picture.altTextTitle = "Math Formula";
          picture.altTextDescription = `Math: ${latexOriginal} #[ID:${eqId}|SVG]`;

          const picRange = picture.getRange();
          if (verticalAlignEx !== 0) {
            picRange.font.position = verticalAlignEx * exToPt;
          } else {
            picRange.font.position = 0;
          }

          const newCc = picRange.insertContentControl();
          newCc.title = "LSTS_Math";
          newCc.tag = eqId + "|SVG";
          newCc.appearance = Word.ContentControlAppearance.hidden;
        } else {
          const insertedRange = range.insertHtml(insertContent, Word.InsertLocation.replace);
          const newCc = insertedRange.insertContentControl();
          newCc.title = "LSTS_Math";
          newCc.tag = eqId + "|MML";
          newCc.appearance = Word.ContentControlAppearance.hidden;
        }
        await context.sync();
      });
    } catch (error) {
      console.error("Lỗi chèn vào Word:", error);
    }
  }

  // ==========================================
  // HOOK TÍNH TOÁN & VẼ SVG XEM TRƯỚC (PREVIEW) - HỖ TRỢ CỘNG/TRỪ/NHÂN/CHIA
  // ==========================================
  const verticalMathPreview = useMemo(() => {
    if (!currentLatex) return null;
    let bieuThuc = currentLatex.trim();

    // 1. Nhận diện loại phép toán
    const laPhepNhan = bieuThuc.includes('\\times') || bieuThuc.includes('\\cdot') || bieuThuc.includes('*') || bieuThuc.includes('x');
    const laPhepChia = bieuThuc.includes('\\frac') || bieuThuc.includes('\\div') || bieuThuc.includes('/');
    const laPhepCong = bieuThuc.includes('+');
    const laPhepTru = bieuThuc.includes('-');

    // Nếu không phải 4 phép toán cơ bản thì không hiển thị Preview
    if (!laPhepNhan && !laPhepChia && !laPhepCong && !laPhepTru) return null;

    const CHAR_W = 16.5;
    const LINE_H = 32;
    const MARGIN = 15;
    const FONT_STYLE = `font-family="Arial, monospace" font-size="28px"`;

    try {
      let svgString = "";
      let mathWidthEx = 0;
      let mathHeightEx = 0;

      if (laPhepNhan) {
        // --- XỬ LÝ NHÂN DỌC ---
        let chuoiSach = bieuThuc.replace(/\\times/g, '*').replace(/\\cdot/g, '*').replace(/x/g, '*').replace(/\\text\{\*\}/g, '*');
        const mangSo = chuoiSach.split('*');
        if (mangSo.length !== 2) return null;

        let strSo1 = mangSo[0].trim();
        let strSo2 = mangSo[1].trim();
        let nSo1 = parseInt(strSo1, 10);
        let nSo2 = parseInt(strSo2, 10);

        if (isNaN(nSo1) || isNaN(nSo2)) return null;

        let ketQuaCuoi = (nSo1 * nSo2).toString();
        const maxChars = Math.max(strSo1.length, strSo2.length, ketQuaCuoi.length);
        const width = maxChars * CHAR_W + MARGIN * 3;

        let linesCount = 2;
        let partials = [];
        if (strSo2.length > 1) {
          for (let i = strSo2.length - 1; i >= 0; i--) {
            partials.push((nSo1 * parseInt(strSo2[i], 10)).toString());
          }
          linesCount += partials.length + 1;
        } else {
          linesCount += 1;
        }

        const height = linesCount * LINE_H + MARGIN * 2;
        mathWidthEx = width / 14;
        mathHeightEx = height / 14;

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
        // --- XỬ LÝ CỘNG / TRỪ DỌC ---
        let toanTu = laPhepCong ? '+' : '-';
        const mangSo = bieuThuc.split(toanTu);
        if (mangSo.length !== 2) return null;

        let strSo1 = mangSo[0].trim();
        let strSo2 = mangSo[1].trim();
        let nSo1 = parseInt(strSo1, 10);
        let nSo2 = parseInt(strSo2, 10);

        if (isNaN(nSo1) || isNaN(nSo2)) return null;

        // Tính kết quả dựa trên dấu
        let ketQuaCuoi = laPhepCong ? (nSo1 + nSo2).toString() : (nSo1 - nSo2).toString();

        const maxChars = Math.max(strSo1.length, strSo2.length, ketQuaCuoi.length);
        const width = maxChars * CHAR_W + MARGIN * 3;
        const height = 3 * LINE_H + MARGIN * 2; // Luôn luôn có 3 dòng

        mathWidthEx = width / 14;
        mathHeightEx = height / 14;

        svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">`;
        const rightEdge = width - MARGIN;

        // Vẽ số thứ 1 và 2 (căn phải)
        svgString += `<text x="${rightEdge}" y="${LINE_H + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo1}</text>`;
        svgString += `<text x="${rightEdge}" y="${LINE_H * 2 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${strSo2}</text>`;

        // Tính toán vị trí dấu cộng/trừ để bám sát lề trái của số dài nhất
        let offsetDauX = rightEdge - Math.max(strSo1.length, strSo2.length) * CHAR_W - 10;
        svgString += `<text x="${offsetDauX}" y="${LINE_H * 1.5 + MARGIN}" text-anchor="end" ${FONT_STYLE} fill="black">${toanTu}</text>`;

        // Gạch ngang ngăn cách
        svgString += `<line x1="${MARGIN}" y1="${LINE_H * 2 + MARGIN + 8}" x2="${rightEdge}" y2="${LINE_H * 2 + MARGIN + 8}" stroke="black" stroke-width="3"/>`;

        // Vẽ kết quả cuối cùng
        let currentY = LINE_H * 3 + MARGIN + 6;
        svgString += `<text x="${rightEdge}" y="${currentY}" text-anchor="end" ${FONT_STYLE} fill="black">${ketQuaCuoi}</text>`;

        svgString += `</svg>`;

      } else {
        // --- XỬ LÝ CHIA DỌC ---
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

        mathWidthEx = width / 14;
        mathHeightEx = height / 14;

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

      // Trả về dữ liệu đóng gói cho Preview
      return { svgString, mathWidthEx, mathHeightEx, bieuThuc };
    } catch (err) {
      return null;
    }
  }, [currentLatex]);

  // ==========================================
  // HÀM 3: CHÈN TRỰC TIẾP DỮ LIỆU ĐÃ TÍNH Ở BƯỚC PREVIEW
  // ==========================================
  const handleInsertVerticalMath = async () => {
    if (!verticalMathPreview) return;
    const { svgString, mathWidthEx, mathHeightEx, bieuThuc } = verticalMathPreview;

    try {
      let utf8Encoded = encodeURIComponent(svgString).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16)));
      let insertContent = window.btoa(utf8Encoded);

      const eqId = "EQV" + new Date().getTime();
      Office.context.document.settings.set(eqId, bieuThuc);
      await Office.context.document.settings.saveAsync();

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("font/size");
        await context.sync();

        const currentFontSize = range.font.size || 12;
        const exToPt = currentFontSize * 0.44;

        const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
        picture.width = mathWidthEx * exToPt;
        picture.height = mathHeightEx * exToPt;
        picture.lockAspectRatio = true;

        // ---> BỔ SUNG: SMART ALT TEXT <---
        picture.altTextTitle = "Math Formula";
        picture.altTextDescription = `Math: ${bieuThuc} #[ID:${eqId}|SVG]`;

        const newCc = picture.getRange().insertContentControl();
        newCc.title = "LSTS_Math";
        newCc.tag = eqId + "|SVG";
        newCc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
      });

    } catch (err) {
      console.error("Lỗi chèn tính dọc: ", err);
    }
  };


  useEffect(() => {
    if (containerRef.current && !mathfieldRef.current) {
      const mf = new MathfieldElement();
      mf.value = "";
      mf.style.width = "100%";
      mf.style.fontSize = "22px";
      mf.style.border = "1px solid #8a8886";
      mf.style.borderRadius = "4px";
      mf.style.padding = "10px";
      mf.style.minHeight = "120px";
      mf.style.backgroundColor = "#ffffff";
      mf.style.boxSizing = "border-box";

      containerRef.current.appendChild(mf);
      mathfieldRef.current = mf;

      mf.inlineShortcuts = {
        ...mf.inlineShortcuts,
        ...customInlineShortcuts,
        ...userDict.inline
      };

      mf.macros = {
        ...mf.macros,
        ...upgradeMacros(customMacros),
        ...upgradeMacros(userDict.macro)
      };

      mf.addEventListener("input", () => {
        setCurrentLatex(mf.getValue("latex-expanded"));
      });
    }

    const selectionHandler = async () => {
      await handleReadEquation();
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      selectionHandler
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: selectionHandler }
      );
    };
  }, [handleReadEquation]);

  const insertSymbol = (latexCode: string) => {
    if (mathfieldRef.current) {
      mathfieldRef.current.insert(latexCode);
      mathfieldRef.current.focus();
    }
  };

  const handleSaveToDict = () => {
    if (!dictKey.trim()) {
      alert("Vui lòng nhập phím tắt!");
      return;
    }
    const updatedDict = { ...userDict };
    updatedDict[dictType][dictKey.trim()] = currentLatex;
    setUserDict(updatedDict);
    localStorage.setItem("math_user_dict", JSON.stringify(updatedDict));

    if (mathfieldRef.current) {
      if (dictType === "inline") {
        mathfieldRef.current.inlineShortcuts = { ...mathfieldRef.current.inlineShortcuts, ...updatedDict.inline };
      } else {
        mathfieldRef.current.macros = { ...mathfieldRef.current.macros, ...upgradeMacros(updatedDict.macro) };
      }
    }
    setIsAddDictOpen(false);
    setDictKey("");
  };

  const handleExportDict = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(userDict, null, 2));
    const downloadNode = document.createElement('a');
    downloadNode.setAttribute("href", dataStr);
    downloadNode.setAttribute("download", "Math_Dictionary.json");
    document.body.appendChild(downloadNode);
    downloadNode.click();
    downloadNode.remove();
  };

  const handleImportDict = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedData = JSON.parse(e.target?.result as string);
        if (importedData.inline && importedData.macro) {
          const mergedDict = {
            inline: { ...importedData.inline, ...userDict.inline },
            macro: { ...importedData.macro, ...userDict.macro }
          };
          setUserDict(mergedDict);
          localStorage.setItem("math_user_dict", JSON.stringify(mergedDict));

          if (mathfieldRef.current) {
            mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...mergedDict.inline };
            mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(mergedDict.macro) };
          }
          alert("Nhập từ điển thành công! Các phím tắt cũ của bạn vẫn được giữ nguyên an toàn.");
        } else {
          alert("Cấu trúc file JSON không đúng chuẩn từ điển!");
        }
      } catch (error) {
        alert("Lỗi đọc file JSON! Hãy kiểm tra lại định dạng file.");
      }
    };
    reader.readAsText(file);
    event.target.value = "";
  };

  const handleOpenEditDict = () => {
    const items: EditItem[] = [];
    Object.entries(userDict.inline).forEach(([k, v]) =>
      items.push({ id: `in_${k}_${Date.now()}`, type: "inline", key: k, value: v })
    );
    Object.entries(userDict.macro).forEach(([k, v]) =>
      items.push({ id: `mc_${k}_${Date.now()}`, type: "macro", key: k, value: v })
    );

    setEditItems(items);
    setSearchTerm("");
    setIsEditDictOpen(true);
  };

  const handleItemChange = (id: string, field: "key" | "value", newValue: string) => {
    setEditItems(prev => prev.map(item =>
      item.id === id ? { ...item, [field]: newValue } : item
    ));
  };

  const handleItemDelete = (id: string) => {
    setEditItems(prev => prev.filter(item => item.id !== id));
  };

  const handleSaveEditedDict = () => {
    const newDict = { inline: {} as Record<string, string>, macro: {} as Record<string, string> };

    editItems.forEach(item => {
      const k = item.key.trim();
      if (k) {
        newDict[item.type][k] = item.value;
      }
    });

    setUserDict(newDict);
    localStorage.setItem("math_user_dict", JSON.stringify(newDict));

    if (mathfieldRef.current) {
      mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...newDict.inline };
      mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(newDict.macro) };
    }
    setIsEditDictOpen(false);
  };

  return (
    <ThemeProvider theme={officeGreenTheme}>
      <div style={{ padding: "8px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8", overflowY: "auto" }}>
        <Stack tokens={{ childrenGap: 12 }}>
          {isOnline && (
            <MessageBar messageBarType={MessageBarType.blocked} isMultiline={true}>
              <strong>Phiên bản Web không được hỗ trợ chèn vào Word!</strong><br />
              Để chèn công thức, vui lòng nhấn nút <b>"Mở trong ứng dụng trên máy tính"</b> (Open in Desktop App) trên thanh công cụ của Word.
            </MessageBar>
          )}

          <div style={{ height: "145px" }}>
            <MathPalette onInsert={insertSymbol} />
          </div>

          <Label style={{ fontSize: "14px", fontWeight: "600" }}>Math Live:</Label>
          <div ref={containerRef} id="mathfield-container" />

          <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: "4px" }}>
              <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>LaTeX:</Label>

              <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                <input type="file" accept=".json" ref={fileInputRef} style={{ display: "none" }} onChange={handleImportDict} />

                <DefaultButton
                  iconProps={{ iconName: 'Upload' }}
                  title="Import Dictionary"
                  onClick={() => fileInputRef.current?.click()}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  title="Export Dictionary"
                  onClick={handleExportDict}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Edit' }}
                  title="Edit Dictionary"
                  onClick={handleOpenEditDict}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Save' }}
                  text="Save to Dict"
                  title="Save to Dictionary"
                  onClick={() => setIsAddDictOpen(true)}
                  disabled={!currentLatex}
                  styles={{
                    root: { height: 24, fontSize: 12, padding: "0 10px", marginLeft: "4px", borderRadius: "6px" },
                    rootDisabled: { backgroundColor: "transparent" },
                    label: { fontWeight: "bold", color: currentLatex ? "#107c41" : undefined },
                    icon: { fontSize: 12, color: currentLatex ? "#107c41" : undefined }
                  }}
                />

                {/* Đã xóa nút Nhân/Chia dọc ở dòng này */}

              </Stack>
            </Stack>

            <textarea
              value={currentLatex}
              onChange={(e) => {
                setCurrentLatex(e.target.value);
                if (mathfieldRef.current) mathfieldRef.current.value = e.target.value;
              }}
              style={{
                width: "100%", minHeight: "60px", marginTop: "4px", padding: "6px", boxSizing: "border-box",
                fontFamily: "Consolas, monospace", fontSize: "13px", color: "#d13438",
                border: "1px solid #ccc", borderRadius: "2px", resize: "vertical"
              }}
            />
          </div>

          <Stack
            horizontal
            horizontalAlign="space-between"
            verticalAlign="center"
            tokens={{ childrenGap: 16 }}
            style={{ marginTop: "12px", marginBottom: "10px" }}
          >
            <Stack tokens={{ childrenGap: 4 }} style={{ flexGrow: 1 }}>
              <Toggle
                label="Mode:"
                inlineLabel
                onText={isUnsupported ? "E-Editor (Only)" : "E-Editor"}
                offText="Office"
                checked={insertMode === "svg"}
                onChange={(_ev, checked) => setInsertMode(checked ? "svg" : "mathml")}
                disabled={isUnsupported}
                styles={{ root: { marginTop: "10px", padding: "10px 0px", marginBottom: isUnsupported ? "4px" : "10px" } }}
              />

              {isUnsupported && (
                <Label style={{ fontSize: "12px", color: "#d13438", marginTop: "0px", marginBottom: "10px", lineHeight: "1.2" }}>
                  ⚠️ Office Math is not supported. Insert E-Editor type only.
                </Label>
              )}
            </Stack>

            <DefaultButton
              primary
              text="Insert"
              onClick={handleInsertEquation}
              style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", fontWeight: "bold", borderRadius: "8px" }}
              title="Insert new equation or update selected equation"
            />
          </Stack>

          {/* ========================================== */}
          {/* KHU VỰC TÍNH TOÁN BẰNG COMPUTE ENGINE      */}
          {/* ========================================== */}
          <div style={{ marginTop: "12px", padding: "12px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd" }}>
            <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", padding: 0 }}>Compute:</Label>

            <Stack horizontal wrap tokens={{ childrenGap: 8 }} style={{ marginBottom: "0px", marginTop: "8px" }}>
              <DefaultButton iconProps={{ iconName: 'Calculator' }} text="Evaluate" onClick={() => handleCompute("evaluate")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Broom' }} text="Simplify" onClick={() => handleCompute("simplify")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'ExploreContent' }} text="Expand" onClick={() => handleCompute("expand")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'MiniExpand' }} text="ExpandAll" onClick={() => handleCompute("expandAll")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Merge' }} text="Factor" onClick={() => handleCompute("factor")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }}/>
              <DefaultButton 
                 iconProps={{ iconName: 'Variable' }} 
                 text="Solve" 
                 onClick={() => handleCompute("solve")} 
                 style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }}
              />
            </Stack>

            {computeResult && (
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: "16px", borderTop: "1px dashed #c8c6c4", paddingTop: "12px" }}>
                <Label style={{ color: "#605e5c", margin: 0 }}>Result:</Label>

                {/* Dùng thẻ math-field chỉ đọc để render công thức kết quả đẹp mắt */}
                <div style={{ padding: "10px", backgroundColor: "white", borderRadius: "4px", fontSize: "22px", minHeight: "50px" }}>
                  <math-field read-only style={{ border: "none", outline: "none", boxShadow: "none" }}>
                    {computeResult}
                  </math-field>
                </div>

                <PrimaryButton
                  text="Insert into Word"
                  iconProps={{ iconName: 'Add' }}
                  onClick={handleInsertComputeResult}
                  style={{ backgroundColor: "#0078d4", borderColor: "transparent", color: "white", width: "fit-content", marginTop: "10px", borderRadius: "6px" }}
                />
              </Stack>
            )}
          </div>

          {/* KHU VỰC PREVIEW PHÉP TÍNH DỌC (SẼ HIỂN THỊ KHI CÓ DATA) */}
          {verticalMathPreview && (
            <div style={{
              marginTop: "4px", padding: "16px", backgroundColor: "#f3f2f1", borderRadius: "6px",
              border: "1px solid #e1dfdd", display: "flex", flexDirection: "column", alignItems: "center"
            }}>
              <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", alignSelf: "flex-start" }}>
                Preview Vertical Math:
              </Label>

              {/* Khung hiển thị SVG */}
              <div
                style={{
                  backgroundColor: "white", padding: "16px", borderRadius: "8px",
                  boxShadow: "0 2px 8px rgba(0,0,0,0.1)", margin: "12px 0"
                }}
                dangerouslySetInnerHTML={{ __html: verticalMathPreview.svgString }}
              />

              <PrimaryButton
                text="Insert into Word"
                iconProps={{ iconName: "CalculatorAddition" }}
                onClick={handleInsertVerticalMath}
                style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }}
              />
            </div>
          )}



          <Dialog hidden={!isAddDictOpen} onDismiss={() => setIsAddDictOpen(false)} dialogContentProps={{ type: DialogType.normal, title: 'Thêm vào từ điển', subText: 'Công thức này sẽ được lưu lại để bạn gõ nhanh lần sau.' }}>
            <ChoiceGroup options={[{ key: 'inline', text: 'Inline (Tự động chuyển đổi khi gõ)' }, { key: 'macro', text: 'Macro (Gõ kèm dấu \\ phía trước để gợi ý)' }]} selectedKey={dictType} onChange={(_, option) => setDictType(option?.key as "inline" | "macro")} />
            <TextField label="Shortcut key (không gõ '\' ở trước)" placeholder="VD: ptb, hpt..." value={dictKey} onChange={(_, val) => setDictKey(val || "")} styles={{ root: { marginTop: 10 }, fieldGroup: { height: 40 }, field: { fontSize: "15px", padding: "8px" } }} />
            <DialogFooter>
              <PrimaryButton onClick={handleSaveToDict} text="Save" />
              <DefaultButton onClick={() => setIsAddDictOpen(false)} text="Cancel" />
            </DialogFooter>
          </Dialog>

          <Panel isOpen={isEditDictOpen} onDismiss={() => setIsEditDictOpen(false)} type={PanelType.medium} headerText="Quản lý Từ điển" closeButtonAriaLabel="Đóng">
            <SearchBox placeholder="Tìm phím tắt hoặc mã LaTeX..." value={searchTerm} onChange={(_, newValue) => setSearchTerm(newValue || "")} styles={{ root: { marginBottom: "16px", marginTop: "10px" } }} />
            <div style={{ maxHeight: "60vh", overflowY: "auto", paddingRight: "8px", overflowX: "hidden" }}>
              {["inline", "macro"].map((type) => {
                const filteredItems = editItems.filter(item => item.type === type && (item.key.toLowerCase().includes(searchTerm.toLowerCase()) || item.value.toLowerCase().includes(searchTerm.toLowerCase())));
                if (filteredItems.length === 0) return null;
                const displayItems = filteredItems.slice(0, 10);
                return (
                  <div key={type} style={{ marginBottom: "20px" }}>
                    <Label style={{ color: "#107c41", borderBottom: "1px solid #c8c6c4", paddingBottom: "4px", marginBottom: "8px" }}> {type === "inline" ? "Phím tắt gõ nhanh (Inline)" : "Lệnh Macro (\\)"} </Label>
                    {displayItems.map((item) => (
                      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }} key={item.id} style={{ marginBottom: "12px" }}>
                        <TextField value={item.key} placeholder="Mã lệnh..." onChange={(_, val) => handleItemChange(item.id, "key", val || "")} styles={{ root: { width: "100px" }, field: { fontWeight: "bold" } }} />
                        <TextField multiline autoAdjustHeight value={item.value} onChange={(_, val) => handleItemChange(item.id, "value", val || "")} styles={{ root: { flexGrow: 1 }, field: { fontFamily: "Consolas, monospace", fontSize: "13px" } }} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title="Xóa phím tắt này" onClick={() => handleItemDelete(item.id)} styles={{ root: { color: "#d13438" }, rootHovered: { backgroundColor: "#fde7e9" } }} />
                      </Stack>
                    ))}
                    {filteredItems.length > 10 && (<Label style={{ fontStyle: "italic", color: "#d13438", fontSize: "12px", marginTop: "4px" }}> * Đang hiển thị 10 / {filteredItems.length} kết quả. Hãy nhập thêm từ khóa để tìm chính xác hơn. </Label>)}
                  </div>
                );
              })}
              {editItems.length === 0 && (<Label style={{ fontStyle: "italic", color: "#a19f9d" }}>Từ điển đang trống.</Label>)}
            </div>
            <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: "20px" }}>
              <PrimaryButton iconProps={{ iconName: 'Save' }} text="Lưu thay đổi" onClick={handleSaveEditedDict} />
              <DefaultButton text="Hủy" onClick={() => setIsEditDictOpen(false)} />
            </Stack>
          </Panel>

        </Stack>
      </div>
    </ThemeProvider>
  );
}