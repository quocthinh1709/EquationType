import * as React from "react";
import { useEffect, useRef, useState, useCallback, useMemo } from "react";
import { DefaultButton, IconButton, Stack, Label, ChoiceGroup, Toggle, MessageBar, MessageBarType } from "@fluentui/react";
import { Dialog, DialogType, DialogFooter, PrimaryButton, TextField, Panel, PanelType, SearchBox, ThemeProvider } from "@fluentui/react";
import { initializeIcons } from '@fluentui/react/lib/Icons';

import { MathfieldElement } from "mathlive";
import { ComputeEngine } from "@cortex-js/compute-engine";
import "mathlive";
import MathPalette from "./MathPalette";
import MathBtns from "./MathPalette";
import { customInlineShortcuts, customMacros } from "./mathShortcuts";
import { unsupportedMathMlCommands } from "./unsupportedCommands";

// --- IMPORT CÁC MODULE ĐÃ TÁCH ---
import { officeGreenTheme, uiText, Language } from "./constants";
import { calculateVariationTableData, createGraphEvaluator } from "./mathEngine";
import { buildVerticalMathPreview, buildVariationTableSvg, buildGraphSvg } from "./svgRenderer";

initializeIcons();
const ce = new ComputeEngine();
MathfieldElement.fontsDirectory = "./fonts";

export default function App() {
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const lastEqIdRef = useRef<string>("");
  const [currentLatex, setCurrentLatex] = useState("");

  const [insertMode, setInsertMode] = useState<string>(() => localStorage.getItem("math_insert_mode") || "mathml");
  const [computeResult, setComputeResult] = useState<string>("");
  const [statusMessage, setStatusMessage] = useState<{ text: string, type: MessageBarType } | null>(null);
  const [language, setLanguage] = useState<Language>(() => (localStorage.getItem("math_user_lang") as Language) || "vi");

  const [verticalMathPreview, setVerticalMathPreview] = useState<any>(null);
  const [variationTablePreview, setVariationTablePreview] = useState<any>(null);

  // --- STATE DÀNH CHO ĐỒ THỊ ---
  const [graphPreview, setGraphPreview] = useState<any>(null);
  const [graphConfig, setGraphConfig] = useState({ xMin: -5, xMax: 5, yMin: -5, yMax: 5 });
  const [graphAnnotations, setGraphAnnotations] = useState<{name: string, x: number, y: number}[]>([]); // Thêm 'name'
  const [tempPointName, setTempPointName] = useState(""); // Lưu tạm chữ 'A'
  const [tempPointX, setTempPointX] = useState("");
  const [graphColor, setGraphColor] = useState("#d13438"); // Lưu màu, mặc định là đỏ sậm
  // Effect tự động vẽ lại Đồ thị mỗi khi cấu hình hoặc biểu thức thay đổi
  useEffect(() => {
    if (!currentLatex || !graphPreview) return;
    try {
      const evaluator = createGraphEvaluator(currentLatex, ce);
      const svgData = buildGraphSvg(evaluator, graphConfig, graphAnnotations, graphColor);
      setGraphPreview({ ...svgData, bieuThuc: currentLatex });
    } catch (e) {
      // Bỏ qua lỗi vẽ đồ thị
    }
  }, [currentLatex, graphConfig, graphAnnotations, graphColor]); // Tự động chạy lại khi xMin, xMax đổi

  // Nút "Vẽ đồ thị"
  const handlePreviewGraph = () => {
    if (!currentLatex) return;
    try {
      const evaluator = createGraphEvaluator(currentLatex, ce);
      const svgData = buildGraphSvg(evaluator, graphConfig, graphAnnotations, graphColor);
      setGraphPreview({ ...svgData, bieuThuc: currentLatex });
    } catch (e) {
      setGraphPreview(null);
    }
  };

  // Nút chèn đồ thị vào Word
  const handleInsertGraph = async () => {
    if (!graphPreview) return;
    const { svgString, mathWidthEx, mathHeightEx, bieuThuc } = graphPreview;
    try {
      let insertContent = window.btoa(encodeURIComponent(svgString).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16))));
      const eqId = "EQG" + new Date().getTime();
      Office.context.document.settings.set(eqId, bieuThuc); await Office.context.document.settings.saveAsync();
      await Word.run(async (context) => {
        const range = context.document.getSelection(); range.load("font/size"); await context.sync();
        const exToPt = (range.font.size || 12) * 0.44;
        const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
        picture.width = mathWidthEx * exToPt; picture.height = mathHeightEx * exToPt; picture.lockAspectRatio = true;
        picture.altTextTitle = "Graph"; picture.altTextDescription = `Math: ${bieuThuc} #[ID:${eqId}|SVG]`;
        const newCc = picture.getRange().insertContentControl();
        newCc.title = "LSTS_Math"; newCc.tag = eqId + "|SVG"; newCc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
      });
    } catch (err) { console.error("Lỗi chèn Đồ thị: ", err); }
  };

  const handleAddAnnotation = () => {
    const xVal = parseFloat(tempPointX);
    if (isNaN(xVal)) return;
    const evaluator = createGraphEvaluator(currentLatex, ce);
    const yVal = evaluator(xVal);
    if (isNaN(yVal)) return;
    setGraphAnnotations([...graphAnnotations, {name: tempPointName.trim(), x: xVal, y: yVal }]);
    setTempPointX("");
    setTempPointName("");
  };



  const [bbtData, setBbtData] = useState<any>(null);

  const defaultInlineRef = useRef<any>(null);
  const defaultMacrosRef = useRef<any>(null);

  const updateBbtCol = (idx: number, field: string, value: string) => {
    if (!bbtData) return;
    const newCols = [...bbtData.cols];
    newCols[idx] = { ...newCols[idx], [field]: value };
    setBbtData({ ...bbtData, cols: newCols });
  };

  const updateBbtLabel = (field: string, value: string) => {
    if (!bbtData) return;
    const currentLabels = bbtData.labels || { x: "x", yPrime: "y'", y: "y" };
    setBbtData({ ...bbtData, labels: { ...currentLabels, [field]: value } });
  };

  // TỰ ĐỘNG CẬP NHẬT SVG KHI DỮ LIỆU BBT THAY ĐỔI
  useEffect(() => {
    if (!bbtData) {
      setVariationTablePreview(null);
      return;
    }
    const svgData = buildVariationTableSvg(bbtData);
    setVariationTablePreview({ ...svgData, bieuThuc: bbtData.bieuThuc });
  }, [bbtData]);

  const t = uiText[language];

  // ==========================================
  // LOGIC COMPUTE ENGINE (Tính toán giá trị)
  // ==========================================
  const handleCompute = (action: "evaluate" | "simplify" | "expand" | "expandAll" | "factor" | "solve" | "derivative" | "integrate" | "limit") => {
    if (!currentLatex) return;
    try {
      const expr = ce.parse(currentLatex);
      let resultLatex = "";

      switch (action) {
        case "evaluate": resultLatex = expr.evaluate().latex; break;
        case "simplify": resultLatex = expr.simplify().latex; break;
        case "expand": resultLatex = ce.expr(["Expand", expr]).evaluate().latex; break;
        case "expandAll": resultLatex = ce.expr(["ExpandAll", expr]).evaluate().latex; break;
        case "factor": resultLatex = ce.expr(["Factor", expr]).evaluate().latex; break;
        case "derivative": resultLatex = ce.expr(["Derivative", expr]).evaluate().latex; break;
        case "integrate":
          let intVar = "x";
          if (expr.unknowns && expr.unknowns.length > 0) intVar = expr.unknowns[0];
          else if (currentLatex.includes("t")) intVar = "t";
          else if (currentLatex.includes("y")) intVar = "y";

          if (currentLatex.includes("\\int")) {
            const evalInt = expr.evaluate();
            if (evalInt && ((evalInt as any).head === "Integrate" || evalInt.latex.includes("\\int"))) {
              resultLatex = t.computeNotSupported;
            } else resultLatex = evalInt.latex;
          } else {
            const evalInt = ce.expr(["Integrate", expr, intVar]).evaluate();
            if (evalInt && ((evalInt as any).head === "Integrate" || evalInt.latex.includes("\\int"))) {
              resultLatex = t.computeNotSupported;
            } else resultLatex = evalInt.latex + " + C";
          }
          break;
        case "limit":
          const limitRegex = /\\lim\s*_\s*\{\s*([a-zA-Z])\s*(?:\\to|\\rightarrow)\s*([^}]+)\s*\}(.*)/;
          const limitMatch = currentLatex.match(limitRegex);
          if (limitMatch) {
            const varName = limitMatch[1];
            let targetLatex = limitMatch[2].trim();
            const funcLatex = limitMatch[3];

            let direction = 0;
            if (targetLatex.includes("^+") || targetLatex.includes("^{+}")) { direction = 1; targetLatex = targetLatex.replace(/\^\{\+\}|\^\+/g, ""); }
            else if (targetLatex.includes("^-") || targetLatex.includes("^{-}")) { direction = -1; targetLatex = targetLatex.replace(/\^\{-\}|\^\-/g, ""); }

            const funcExpr = ce.parse(funcLatex);
            const targetExpr = ce.parse(targetLatex);

            try {
              let testValueStr = targetLatex;
              if (targetLatex.includes("\\infty")) testValueStr = targetLatex.includes("-") ? "-10000000" : "10000000";
              else {
                let numTarget = parseFloat((targetExpr as any).numericValue || targetExpr.value || targetLatex);
                if (!isNaN(numTarget)) {
                  if (direction === 1) numTarget += 1e-7; else if (direction === -1) numTarget -= 1e-7; else numTarget += 1e-8;
                  testValueStr = numTarget.toString();
                }
              }

              const numericValObj = funcExpr.subs({ [varName]: ce.parse(testValueStr) }).N();
              let numericRes = NaN;
              if (numericValObj && typeof (numericValObj as any).numericValue === 'number') numericRes = (numericValObj as any).numericValue;
              else numericRes = parseFloat(numericValObj.toString());

              if (!isNaN(numericRes)) {
                if (numericRes > 1e5 || !isFinite(numericRes)) resultLatex = "+\\infty";
                else if (numericRes < -1e5) resultLatex = "-\\infty";
                else resultLatex = (Math.round(numericRes * 10000) / 10000).toString();
              } else {
                let evalLimit = direction !== 0 ? ce.expr(["Limit", funcExpr, varName, targetExpr, direction]).evaluate() : ce.expr(["Limit", funcExpr, varName, targetExpr]).evaluate();
                resultLatex = (evalLimit && ((evalLimit as any).head === "Limit" || evalLimit.latex.includes("\\lim"))) ? t.computeNotSupported : evalLimit.latex;
              }
            } catch { resultLatex = t.computeNotSupported; }
          } else {
            const evalResult = expr.evaluate();
            resultLatex = (evalResult && ((evalResult as any).head === "Limit" || evalResult.latex.includes("\\lim"))) ? t.computeNotSupported : (evalResult ? evalResult.latex : t.computeNotSupported);
          }
          break;
        case "solve":
          let vars: string[] = (expr.unknowns && expr.unknowns.length > 0) ? [...expr.unknowns] : [];
          if (vars.length === 0) {
            if (currentLatex.includes("x")) vars.push("x");
            if (currentLatex.includes("y")) vars.push("y");
            if (currentLatex.includes("z")) vars.push("z");
            if (vars.length === 0) vars = ["x"];
          }
          if (vars.length > 1) {
            if (typeof expr.solve === "function") {
              const solution = expr.solve(vars) as Record<string, any>;
              if (solution && typeof solution === "object" && !Array.isArray(solution) && Object.keys(solution).length > 0) {
                let casesContent = vars.map(v => solution[v] ? `${v} = ${solution[v].latex}` : "").filter(Boolean).join(" \\\\ ");
                resultLatex = `\\displaystyle \\begin{cases} ${casesContent} \\end{cases}`;
              } else resultLatex = t.computeSystemNoSolution;
            } else resultLatex = t.computeNotSupported;
          } else {
            const v = vars[0];
            const solveResult = ce.expr(["Solve", expr, v]).evaluate();
            if (solveResult && solveResult.latex.includes("Solve")) resultLatex = t.computeSolveNotSupported;
            else {
              resultLatex = `${v} \\in \\{` + solveResult.latex + "\\}";
              resultLatex = resultLatex.replace(/\\\{\\\{/g, "\\{").replace(/\\\}\\\}/g, "\\}");
            }
          }
          break;
      }
      setComputeResult(resultLatex);
    } catch (error) { setComputeResult("t.computeError"); }
  };

  const handleInsertComputeResult = () => {
    if (!computeResult || !mathfieldRef.current) return;
    mathfieldRef.current.value = computeResult;
    setCurrentLatex(computeResult);
    setTimeout(() => { handleInsertEquation(); }, 100);
  };

  // TỪ ĐIỂN
  const [userDict, setUserDict] = useState<{ inline: Record<string, string>, macro: Record<string, string> }>(() => {
    try { const saved = localStorage.getItem("math_user_dict"); return saved ? JSON.parse(saved) : { inline: {}, macro: {} }; } catch { return { inline: {}, macro: {} }; }
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
      localStorage.setItem("math_insert_mode", "svg");
    }
  }, [isUnsupported, insertMode]);

  const upgradeMacros = (macroDict: Record<string, any>) => {
    const upgraded: Record<string, any> = {};
    Object.keys(macroDict).forEach(key => {
      const val = macroDict[key];
      if (typeof val === "string") upgraded[key] = { def: val, captureSelection: false, expand: true };
      else upgraded[key] = val;
    });
    return upgraded;
  };

  // OFFICE API INTEGRATION
  const handleReadEquation = useCallback(async () => {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const cc = range.parentContentControlOrNullObject; cc.load("title, tag");
        const ccs = range.contentControls; ccs.load("items/title, items/tag");
        const pics = range.inlinePictures; pics.load("items/altTextDescription");
        await context.sync();

        let newFullTag = "";
        let newTargetCC: Word.ContentControl | null = null;

        if (ccs.items.length > 0 && ccs.items[0].title === "LSTS_Math") { newFullTag = ccs.items[0].tag; newTargetCC = ccs.items[0]; }
        else if (!cc.isNullObject && cc.title === "LSTS_Math") { newFullTag = cc.tag; newTargetCC = cc; }
        else if (pics.items.length > 0 && pics.items[0].altTextDescription && pics.items[0].altTextDescription.includes("#[ID:")) {
          const match = pics.items[0].altTextDescription.match(/#\[ID:(.*?)\]/);
          if (match && match[1]) newFullTag = match[1];
        }

        const newEqId = newFullTag ? newFullTag.split("|")[0] : "";
        const newEqType = newFullTag ? (newFullTag.split("|")[1] || "SVG") : "";

        if (lastEqIdRef.current && lastEqIdRef.current !== newFullTag) {
          const oldControls = context.document.contentControls.getByTag(lastEqIdRef.current); oldControls.load("items"); await context.sync();
          if (oldControls.items.length > 0) {
            const oldCC = oldControls.items[0]; const oldRange = oldCC.getRange(); const oldPics = oldRange.inlinePictures; oldPics.load("items"); oldRange.load("text"); await context.sync();
            const oldEqType = lastEqIdRef.current.split("|")[1] || "SVG";
            const rawText = oldRange.text; const cleanText = rawText.replace(/[\s\u200B-\u200D\uFEFF\uFFFC]/g, '').toLowerCase();
            let isOldEmptied = false;
            if (oldEqType === "SVG") { if (oldPics.items.length === 0) isOldEmptied = true; } else { if (rawText === "" || cleanText.includes("clickortap") || cleanText.includes("vănbản")) isOldEmptied = true; }
            if (isOldEmptied) { oldCC.delete(false); await context.sync(); }
          }
        }

        if (newFullTag) {
          lastEqIdRef.current = newFullTag; let isEmptied = false;
          if (newTargetCC) {
            const ccRange = newTargetCC.getRange(); const ccPics = ccRange.inlinePictures; ccPics.load("items"); ccRange.load("text"); await context.sync();
            const cleanText = ccRange.text.replace(/[\s\u200B-\u200D\uFEFF\uFFFC]/g, '').toLowerCase();
            isEmptied = newEqType === "SVG" ? ccPics.items.length === 0 : (ccRange.text === "" || cleanText.includes("clickortap") || cleanText.includes("vănbản"));
          }
          if (isEmptied && newTargetCC) {
            if (mathfieldRef.current) mathfieldRef.current.value = "";
            setCurrentLatex(""); lastEqIdRef.current = ""; newTargetCC.delete(false); await context.sync();
          } else {
            const latexCode = Office.context.document.settings.get(newEqId);
            if (latexCode && mathfieldRef.current && mathfieldRef.current.value !== latexCode) { mathfieldRef.current.value = latexCode; setCurrentLatex(latexCode); }
          }
        } else lastEqIdRef.current = "";
      });
    } catch (error) { console.error("Lỗi đọc công thức:", error); }
  }, []);

  const handleInsertEquation = async () => {
    if (!mathfieldRef.current) return;
    const latexOriginal = mathfieldRef.current.value;
    const latexExpanded = mathfieldRef.current.getValue("latex-expanded");

    let insertContent = ""; let isInsertingImage = false; let mathWidthEx = 0; let mathHeightEx = 0; let verticalAlignEx = 0;

    if (insertMode === "mathml") {
      const hiddenDiv = document.createElement("div"); hiddenDiv.style.display = "none"; document.body.appendChild(hiddenDiv);
      const tempMf = new MathfieldElement(); hiddenDiv.appendChild(tempMf); tempMf.value = latexExpanded;
      const mathML = tempMf.getValue("math-ml" as any); document.body.removeChild(hiddenDiv);
      insertContent = mathML.includes("<math") ? mathML : `<math xmlns="http://www.w3.org/1998/Math/MathML">${mathML}</math>`;
    } else {
      try {
        const mathJaxObj = (window as any).MathJax;
        if (mathJaxObj && mathJaxObj.tex2svg) {
          const svgNode = mathJaxObj.tex2svg(latexExpanded); const svgElement = svgNode.querySelector("svg");
          if (svgElement) {
            mathWidthEx = parseFloat(svgElement.getAttribute("width") || "0"); mathHeightEx = parseFloat(svgElement.getAttribute("height") || "0");
            const widthPt = mathWidthEx * 6; const heightPt = mathHeightEx * 6;
            const vaMatch = (svgElement.getAttribute("style") || "").match(/vertical-align:\s*([-0-9.]+)ex/);
            if (vaMatch && vaMatch[1]) verticalAlignEx = parseFloat(vaMatch[1]);
            svgElement.setAttribute("width", widthPt + "pt"); svgElement.setAttribute("height", heightPt + "pt");
            svgElement.setAttribute("xmlns", "http://www.w3.org/2000/svg"); svgElement.setAttribute("shape-rendering", "geometricPrecision");
            let svgData = new XMLSerializer().serializeToString(svgElement).replace(/currentColor/gi, "#000000");
            insertContent = window.btoa(encodeURIComponent(svgData).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16))));
            isInsertingImage = true;
          }
        }
      } catch (error) { console.error("Lỗi vẽ PNG:", error); }
    }

    try {
      const eqId = "EQ" + new Date().getTime() + "_" + Math.random().toString(36).substring(2, 6);
      Office.context.document.settings.set(eqId, latexExpanded); await Office.context.document.settings.saveAsync();
      await Word.run(async (context) => {
        const range = context.document.getSelection(); range.load("font/size"); await context.sync();
        const currentFontSize = range.font.size || 12;
        if (isInsertingImage) {
          if (!insertContent) return;
          const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
          const exToPt = currentFontSize * 0.44;
          if (mathWidthEx > 0 && mathHeightEx > 0) { picture.width = mathWidthEx * exToPt; picture.height = mathHeightEx * exToPt; }
          picture.lockAspectRatio = true;
          picture.altTextTitle = "Math Formula"; picture.altTextDescription = `Math: ${latexOriginal} #[ID:${eqId}|SVG]`;
          const picRange = picture.getRange();
          picRange.font.position = verticalAlignEx !== 0 ? verticalAlignEx * exToPt : 0;
          const newCc = picRange.insertContentControl();
          newCc.title = "LSTS_Math"; newCc.tag = eqId + "|SVG"; newCc.appearance = Word.ContentControlAppearance.hidden;
        } else {
          const insertedRange = range.insertHtml(insertContent, Word.InsertLocation.replace);
          const newCc = insertedRange.insertContentControl();
          newCc.title = "LSTS_Math"; newCc.tag = eqId + "|MML"; newCc.appearance = Word.ContentControlAppearance.hidden;
        }
        await context.sync();
      });
    } catch (error) { console.error("Lỗi chèn vào Word:", error); }
  }

  // --- HÀM TẠO PREVIEW ---
  const handlePreviewVertical = () => {
    if (!currentLatex) { setVerticalMathPreview(null); return; }
    let bieuThuc = currentLatex.trim();
    if (bieuThuc.includes('y=') || bieuThuc.includes('f(x)=')) { setVerticalMathPreview(null); return; }

    const preview = buildVerticalMathPreview(bieuThuc);
    if (preview) setVerticalMathPreview(preview);
    else {
      setStatusMessage({ text: "Vui lòng nhập phép tính +, -, *, / cơ bản để xem tính dọc.", type: MessageBarType.warning });
      setVerticalMathPreview(null);
    }
  };

  const handlePreviewVariation = () => {
    if (!currentLatex) { setVariationTablePreview(null); return; }
    const latex = currentLatex.trim();
    try {
      const cols = calculateVariationTableData(latex, ce);
      setBbtData({ cols, bieuThuc: latex, labels: { x: "x", yPrime: "y'", y: "y" } });
    } catch (err) {
      console.error("Lỗi vẽ BBT:", err);
      setVariationTablePreview(null);
    }
  };

  const handleInsertVerticalMath = async () => {
    if (!verticalMathPreview) return;
    const { svgString, mathWidthEx, mathHeightEx, bieuThuc } = verticalMathPreview;
    try {
      let insertContent = window.btoa(encodeURIComponent(svgString).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16))));
      const eqId = "EQV" + new Date().getTime();
      Office.context.document.settings.set(eqId, bieuThuc); await Office.context.document.settings.saveAsync();
      await Word.run(async (context) => {
        const range = context.document.getSelection(); range.load("font/size"); await context.sync();
        const exToPt = (range.font.size || 12) * 0.44;
        const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
        picture.width = mathWidthEx * exToPt; picture.height = mathHeightEx * exToPt; picture.lockAspectRatio = true;
        picture.altTextTitle = "Math Formula"; picture.altTextDescription = `Math: ${bieuThuc} #[ID:${eqId}|SVG]`;
        const newCc = picture.getRange().insertContentControl();
        newCc.title = "LSTS_Math"; newCc.tag = eqId + "|SVG"; newCc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
      });
    } catch (err) { console.error("Lỗi chèn tính dọc: ", err); }
  };

  const handleInsertVariationTable = async () => {
    if (!variationTablePreview) return;
    const { svgString, mathWidthEx, mathHeightEx, bieuThuc } = variationTablePreview;
    try {
      let insertContent = window.btoa(encodeURIComponent(svgString).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16))));
      const eqId = "EQT" + new Date().getTime();
      Office.context.document.settings.set(eqId, bieuThuc); await Office.context.document.settings.saveAsync();
      await Word.run(async (context) => {
        const range = context.document.getSelection(); range.load("font/size"); await context.sync();
        const exToPt = (range.font.size || 12) * 0.44;
        const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
        picture.width = mathWidthEx * exToPt; picture.height = mathHeightEx * exToPt; picture.lockAspectRatio = true;
        picture.altTextTitle = "Variation Table"; picture.altTextDescription = `Math: ${bieuThuc} #[ID:${eqId}|SVG]`;
        const newCc = picture.getRange().insertContentControl();
        newCc.title = "LSTS_Math"; newCc.tag = eqId + "|SVG"; newCc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
      });
    } catch (err) { console.error("Lỗi chèn BBT: ", err); }
  };

  useEffect(() => {
    if (containerRef.current && !mathfieldRef.current) {
      const mf = new MathfieldElement(); mf.value = ""; mf.style.width = "100%"; mf.style.fontSize = "22px";
      mf.style.border = "1px solid #8a8886"; mf.style.borderRadius = "4px"; mf.style.padding = "10px";
      mf.style.minHeight = "120px"; mf.style.backgroundColor = "#ffffff"; mf.style.boxSizing = "border-box";
      containerRef.current.appendChild(mf); mathfieldRef.current = mf;
      mf.addEventListener("input", () => {
        setCurrentLatex(mf.getValue("latex-expanded"));
        setVerticalMathPreview(null); setVariationTablePreview(null); setBbtData(null);
      });
    }
    const selectionHandler = async () => { await handleReadEquation(); };
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, selectionHandler);
    return () => { Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: selectionHandler }); };
  }, [handleReadEquation]);

  useEffect(() => {
    if (!mathfieldRef.current) return;
    const mf = mathfieldRef.current;
    const baseInlines = { ...defaultInlineRef.current }; const baseMacros = { ...defaultMacrosRef.current };
    const userInlines = { ...userDict.inline }; const userMacros = { ...upgradeMacros(userDict.macro) };

    if (language === "vi") {
      mf.inlineShortcuts = { ...baseInlines, ...customInlineShortcuts, ...userInlines };
      mf.macros = { ...baseMacros, ...upgradeMacros(customMacros), ...userMacros };
    } else {
      mf.inlineShortcuts = { ...baseInlines, ...userInlines };
      mf.macros = { ...baseMacros, ...userMacros };
    }
  }, [language, userDict]);

  const insertSymbol = (latexCode: string) => { if (mathfieldRef.current) { mathfieldRef.current.insert(latexCode); mathfieldRef.current.focus(); } };

  const handleSaveToDict = () => {
    if (!dictKey.trim()) { setStatusMessage({ text: t.enterShortcutAlert, type: MessageBarType.warning }); return; }
    const updatedDict = { ...userDict }; updatedDict[dictType][dictKey.trim()] = currentLatex;
    setUserDict(updatedDict); localStorage.setItem("math_user_dict", JSON.stringify(updatedDict));
    setIsAddDictOpen(false); setDictKey("");
  };

  const handleExportDict = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(userDict, null, 2));
    const downloadNode = document.createElement('a'); downloadNode.setAttribute("href", dataStr); downloadNode.setAttribute("download", "Math_Dictionary.json");
    document.body.appendChild(downloadNode); downloadNode.click(); downloadNode.remove();
  };

  const handleImportDict = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedData = JSON.parse(e.target?.result as string);
        if (importedData.inline && importedData.macro) {
          const mergedDict = { inline: { ...importedData.inline, ...userDict.inline }, macro: { ...importedData.macro, ...userDict.macro } };
          setUserDict(mergedDict); localStorage.setItem("math_user_dict", JSON.stringify(mergedDict));
          setTimeout(() => { setStatusMessage({ text: t.importSuccessAlert, type: MessageBarType.success }); }, 100);
        } else setStatusMessage({ text: t.importStructErrorAlert, type: MessageBarType.error });
      } catch (error) { setStatusMessage({ text: t.importReadErrorAlert, type: MessageBarType.error }); }
    };
    reader.readAsText(file); event.target.value = "";
  };

  const handleOpenEditDict = () => {
    const items: EditItem[] = [];
    Object.entries(userDict.inline).forEach(([k, v]) => items.push({ id: `in_${k}_${Date.now()}`, type: "inline", key: k, value: v }));
    Object.entries(userDict.macro).forEach(([k, v]) => items.push({ id: `mc_${k}_${Date.now()}`, type: "macro", key: k, value: v }));
    setEditItems(items); setSearchTerm(""); setIsEditDictOpen(true);
  };

  const handleItemChange = (id: string, field: "key" | "value", newValue: string) => { setEditItems(prev => prev.map(item => item.id === id ? { ...item, [field]: newValue } : item)); };
  const handleItemDelete = (id: string) => { setEditItems(prev => prev.filter(item => item.id !== id)); };
  const handleSaveEditedDict = () => {
    const newDict = { inline: {} as Record<string, string>, macro: {} as Record<string, string> };
    editItems.forEach(item => { const k = item.key.trim(); if (k) { newDict[item.type][k] = item.value; } });
    setUserDict(newDict); localStorage.setItem("math_user_dict", JSON.stringify(newDict)); setIsEditDictOpen(false);
  };

  const isVerticalMathReady = useMemo(() => {
    if (!currentLatex) return false;
    const bieuThuc = currentLatex.trim();
    if (bieuThuc.includes('y=') || bieuThuc.includes('f(x)=')) return false;
    return bieuThuc.includes('\\times') || bieuThuc.includes('\\cdot') || bieuThuc.includes('*') || bieuThuc.includes('x') || bieuThuc.includes('\\frac') || bieuThuc.includes('\\div') || bieuThuc.includes('/') || bieuThuc.includes('+') || bieuThuc.includes('-');
  }, [currentLatex]);

  const isBBTReady = useMemo(() => {
    if (!currentLatex) return false;
    const bieuThuc = currentLatex.toLowerCase();
    if (!bieuThuc.includes('x')) return false;
    if (['\\sin', '\\cos', '\\tan', '\\cot', '\\ln', '\\log', '\\sqrt', '|', '\\abs'].some(keyword => bieuThuc.includes(keyword))) return false;
    if (/\^(x|\{[^}]*x[^}]*\})/.test(bieuThuc)) return false;
    return true;
  }, [currentLatex]);

  return (
    <ThemeProvider theme={officeGreenTheme}>
      <div style={{ padding: "8px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" }}>
          <Label style={{ fontSize: "14px", fontWeight: "bold", margin: 0, color: "#107c41" }}>Easy Efficient Equation Editor</Label>
          <Toggle onText="Tiếng Việt" offText="English" checked={language === 'vi'} onChange={(_, checked) => { const newLang = checked ? 'vi' : 'en'; setLanguage(newLang); localStorage.setItem("math_user_lang", newLang); }} styles={{ root: { margin: 0 } }} />
        </div>
        <Stack tokens={{ childrenGap: 12 }}>
          {isOnline && <MessageBar messageBarType={MessageBarType.blocked} isMultiline={true}><span style={{ whiteSpace: "pre-line" }}>{t.webWarning}</span></MessageBar>}
          {statusMessage && <MessageBar messageBarType={statusMessage.type} onDismiss={() => setStatusMessage(null)} isMultiline={false}>{statusMessage.text}</MessageBar>}

          <div style={{ height: "145px" }}><MathPalette onInsert={insertSymbol} language={language} /></div>

          <Label style={{ fontSize: "14px", fontWeight: "600" }}>{t.mathLiveLabel}</Label>
          <div ref={containerRef} id="mathfield-container" />

          <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: "4px" }}>
              <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>{t.latexLabel}</Label>
              <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                <input type="file" accept=".json" ref={fileInputRef} style={{ display: "none" }} onChange={handleImportDict} />
                <DefaultButton iconProps={{ iconName: 'Upload' }} title={t.importDict} onClick={() => fileInputRef.current?.click()} styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }} />
                <DefaultButton iconProps={{ iconName: 'Download' }} title={t.exportDict} onClick={handleExportDict} styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }} />
                <DefaultButton iconProps={{ iconName: 'Edit' }} title={t.editDict} onClick={handleOpenEditDict} styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }} />
                <DefaultButton iconProps={{ iconName: 'Save' }} text="Save to Dict" title={t.saveToDict} onClick={() => setIsAddDictOpen(true)} disabled={!currentLatex} styles={{ root: { height: 24, fontSize: 12, padding: "0 10px", marginLeft: "4px", borderRadius: "6px" }, rootDisabled: { backgroundColor: "transparent" }, label: { fontWeight: "bold", color: currentLatex ? "#107c41" : undefined }, icon: { fontSize: 12, color: currentLatex ? "#107c41" : undefined } }} />
              </Stack>
            </Stack>
            <textarea value={currentLatex} onChange={(e) => { setCurrentLatex(e.target.value); if (mathfieldRef.current) mathfieldRef.current.value = e.target.value; }} style={{ width: "100%", minHeight: "60px", marginTop: "4px", padding: "6px", boxSizing: "border-box", fontFamily: "Consolas, monospace", fontSize: "13px", color: "#d13438", border: "1px solid #ccc", borderRadius: "2px", resize: "vertical" }} />
          </div>

          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 16 }} style={{ marginTop: "12px", marginBottom: "10px" }}>
            <Stack tokens={{ childrenGap: 4 }} style={{ flexGrow: 1 }}>
              <Toggle label={t.modeLabel} inlineLabel onText={isUnsupported ? t.eEditorOnly : t.eEditor} offText={t.office} checked={insertMode === "svg"} onChange={(_ev, checked) => { const newMode = checked ? "svg" : "mathml"; setInsertMode(newMode); localStorage.setItem("math_insert_mode", newMode); }} disabled={isUnsupported} styles={{ root: { marginTop: "10px", padding: "10px 0px", marginBottom: isUnsupported ? "4px" : "10px" } }} />
              {isUnsupported && <Label style={{ fontSize: "12px", color: "#d13438", marginTop: "0px", marginBottom: "10px", lineHeight: "1.2" }}>{t.officeMathWarning}</Label>}
            </Stack>
            <DefaultButton primary text={t.insertBtn} onClick={handleInsertEquation} style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", fontWeight: "bold", borderRadius: "8px" }} title="Insert new equation or update selected equation" />
          </Stack>

          <div style={{ marginTop: "12px", padding: "12px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd" }}>
            <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", padding: 0 }}>{t.computeLabel}</Label>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }} style={{ marginBottom: "0px", marginTop: "8px" }}>
              <DefaultButton iconProps={{ iconName: 'Calculator' }} text={t.evaluate} onClick={() => handleCompute("evaluate")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Broom' }} text={t.simplify} onClick={() => handleCompute("simplify")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'ExploreContent' }} text={t.expand} onClick={() => handleCompute("expand")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'MiniExpand' }} text={t.expandAll} onClick={() => handleCompute("expandAll")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Merge' }} text={t.factor} onClick={() => handleCompute("factor")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Variable' }} text={t.solve} onClick={() => handleCompute("solve")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'CalculatorSubtract' }} text={t.derivative} onClick={() => handleCompute("derivative")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'CalculatorAddition' }} text={t.integrate} onClick={() => handleCompute("integrate")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'CalculatorMultiply' }} text={t.limit} onClick={() => handleCompute("limit")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'NumberField' }} text="Tính dọc" onClick={handlePreviewVertical} disabled={!isVerticalMathReady} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Table' }} text="Vẽ BBT" onClick={handlePreviewVariation} disabled={!isBBTReady} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'LineChart' }} text="Vẽ Đồ thị" onClick={handlePreviewGraph} disabled={!isBBTReady} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
            </Stack>

            {computeResult && (
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: "16px", borderTop: "1px dashed #c8c6c4", paddingTop: "12px" }}>
                <Label style={{ color: "#605e5c", margin: 0 }}>{t.resultLabel}</Label>
                <div style={{ padding: "10px", backgroundColor: "white", borderRadius: "4px", fontSize: "22px", minHeight: "50px" }}>
                  <math-field read-only style={{ border: "none", outline: "none", boxShadow: "none" }}>{computeResult}</math-field>
                </div>
                <PrimaryButton text={t.insertToWord} iconProps={{ iconName: 'Add' }} onClick={handleInsertComputeResult} style={{ backgroundColor: "#0078d4", borderColor: "transparent", color: "white", width: "fit-content", marginTop: "10px", borderRadius: "6px" }} />
              </Stack>
            )}
          </div>

          {verticalMathPreview && (
            <div style={{ marginTop: "4px", padding: "16px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd", display: "flex", flexDirection: "column", alignItems: "center" }}>
              <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", alignSelf: "flex-start" }}>{t.previewVertical}</Label>
              <div style={{ backgroundColor: "white", padding: "16px", borderRadius: "8px", boxShadow: "0 2px 8px rgba(0,0,0,0.1)", margin: "12px 0" }} dangerouslySetInnerHTML={{ __html: verticalMathPreview.svgString }} />
              <PrimaryButton text={t.insertToWord} iconProps={{ iconName: "CalculatorAddition" }} onClick={handleInsertVerticalMath} style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }} />
            </div>
          )}

          {variationTablePreview && bbtData && (
            <div style={{ marginTop: "4px", padding: "16px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd", display: "flex", flexDirection: "column", alignItems: "center" }}>
              <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", alignSelf: "flex-start" }}>{t.previewVariation}</Label>
              <div style={{ backgroundColor: "white", padding: "10px", borderRadius: "8px", width: "100%", boxShadow: "0 2px 8px rgba(0,0,0,0.1)", margin: "12px 0", overflowX: "auto", textAlign: 'center' }} dangerouslySetInnerHTML={{ __html: variationTablePreview.svgString }} />

              <div style={{ width: "100%", marginBottom: "16px", backgroundColor: "#ffffff", padding: "12px", borderRadius: "8px", border: "1px solid #d1d1d1", overflowX: "auto" }}>
                <Label style={{ fontSize: "12px", color: "#605e5c", marginBottom: "8px" }}>Bảng chỉnh sửa thủ công (nếu cần):</Label>
                <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: '13px' }}>
                  <tbody>
                    <tr>
                      <td style={{ paddingRight: '12px', textAlign: 'center', width: '50px' }}><input type="text" value={bbtData.labels?.x || 'x'} onChange={(e) => updateBbtLabel('x', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px', fontWeight: 'bold', fontStyle: 'italic', backgroundColor: '#e1dfdd' }} /></td>
                      {bbtData.cols.map((col: any, idx: number) => (
                        <td key={`x-${idx}`} style={{ padding: '2px' }}>{col.type === 'val' && <input type="text" value={col.x} onChange={(e) => updateBbtCol(idx, 'x', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px' }} />}</td>
                      ))}
                    </tr>
                    <tr>
                      <td style={{ paddingRight: '12px', textAlign: 'center' }}><input type="text" value={bbtData.labels?.yPrime || "y'"} onChange={(e) => updateBbtLabel('yPrime', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px', fontWeight: 'bold', fontStyle: 'italic', backgroundColor: '#e1dfdd' }} /></td>
                      {bbtData.cols.map((col: any, idx: number) => (
                        <td key={`yp-${idx}`} style={{ padding: '2px' }}>{col.type === 'interval' ? <input type="text" value={col.sign} onChange={(e) => updateBbtCol(idx, 'sign', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', color: 'red', fontWeight: 'bold', border: '1px solid #ccc', borderRadius: '4px', padding: '4px' }} /> : <input type="text" value={col.yPrime !== undefined ? col.yPrime : ''} onChange={(e) => updateBbtCol(idx, 'yPrime', e.target.value)} disabled={col.isDis} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px', backgroundColor: col.isDis ? '#f0f0f0' : 'white' }} />}</td>
                      ))}
                    </tr>
                    <tr>
                      <td style={{ paddingRight: '12px', textAlign: 'center' }}><input type="text" value={bbtData.labels?.y || 'y'} onChange={(e) => updateBbtLabel('y', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px', fontWeight: 'bold', fontStyle: 'italic', backgroundColor: '#e1dfdd' }} /></td>
                      {bbtData.cols.map((col: any, idx: number) => (
                        <td key={`y-${idx}`} style={{ padding: '2px' }}>{col.type === 'val' && (col.isDis ? <div style={{ display: 'flex', gap: '4px' }}><input type="text" value={col.yLeft} onChange={(e) => updateBbtCol(idx, 'yLeft', e.target.value)} title="Giới hạn trái" style={{ width: '50%', minWidth: '30px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px' }} /><input type="text" value={col.yRight} onChange={(e) => updateBbtCol(idx, 'yRight', e.target.value)} title="Giới hạn phải" style={{ width: '50%', minWidth: '30px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px' }} /></div> : <input type="text" value={col.y} onChange={(e) => updateBbtCol(idx, 'y', e.target.value)} style={{ width: '100%', minWidth: '40px', textAlign: 'center', border: '1px solid #ccc', borderRadius: '4px', padding: '4px' }} />)}</td>
                      ))}
                    </tr>
                  </tbody>
                </table>
              </div>
              <PrimaryButton text={t.insertToWord} iconProps={{ iconName: "Table" }} onClick={handleInsertVariationTable} style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }} />
            </div>
          )}


          {/* GIAO DIỆN XEM TRƯỚC VÀ CÀI ĐẶT ĐỒ THỊ */}
          {graphPreview && (
            <div style={{ marginTop: "4px", padding: "16px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd", display: "flex", flexDirection: "column", alignItems: "center" }}>
              <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", alignSelf: "flex-start" }}>Bản xem trước Đồ thị:</Label>

              <div style={{ backgroundColor: "white", padding: "10px", borderRadius: "8px", width: "100%", boxShadow: "0 2px 8px rgba(0,0,0,0.1)", margin: "12px 0", overflowX: "auto", textAlign: 'center' }} dangerouslySetInnerHTML={{ __html: graphPreview.svgString }} />

              {/* BẢNG CÀI ĐẶT TRỤC OXY VÀ MÀU SẮC */}
              <div style={{ width: "100%", marginBottom: "16px", backgroundColor: "#ffffff", padding: "12px", borderRadius: "8px", border: "1px solid #d1d1d1" }}>
                
                {/* Hàng 1: Tiêu đề và Nút chọn màu */}
                <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: "8px" }}>
                   <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>Giới hạn trục tọa độ:</Label>
                   <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>Màu đồ thị:</Label>
                      <input 
                          type="color" 
                          value={graphColor} 
                          onChange={(e) => setGraphColor(e.target.value)} 
                          style={{ width: "24px", height: "24px", padding: "0", border: "1px solid #ccc", cursor: "pointer", borderRadius: "4px" }} 
                      />
                   </Stack>
                </Stack>
                
                {/* Hàng 2: Các ô nhập giới hạn X, Y */}
                <Stack horizontal tokens={{ childrenGap: 10 }} verticalAlign="end" wrap>
                   <TextField label="X Tối thiểu" type="number" value={graphConfig.xMin.toString()} onChange={(_, v) => setGraphConfig({...graphConfig, xMin: parseFloat(v || "-5")})} styles={{ root: { width: 80 } }} />
                   <TextField label="X Tối đa" type="number" value={graphConfig.xMax.toString()} onChange={(_, v) => setGraphConfig({...graphConfig, xMax: parseFloat(v || "5")})} styles={{ root: { width: 80 } }} />
                   <TextField label="Y Tối thiểu" type="number" value={graphConfig.yMin.toString()} onChange={(_, v) => setGraphConfig({...graphConfig, yMin: parseFloat(v || "-5")})} styles={{ root: { width: 80 } }} />
                   <TextField label="Y Tối đa" type="number" value={graphConfig.yMax.toString()} onChange={(_, v) => setGraphConfig({...graphConfig, yMax: parseFloat(v || "5")})} styles={{ root: { width: 80 } }} />
                </Stack>

                {/* Hàng 3: Gắn tên điểm và tọa độ */}
                <Label style={{ fontSize: "12px", color: "#605e5c", marginTop: "16px", marginBottom: "8px" }}>Gắn tọa độ điểm (Ví dụ: Cực trị):</Label>
                <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                   <TextField placeholder="Tên điểm (VD: A)" value={tempPointName} onChange={(_, v) => setTempPointName(v || "")} styles={{ root: { width: 120 } }} />
                   <TextField placeholder="Nhập giá trị x..." value={tempPointX} onChange={(_, v) => setTempPointX(v || "")} styles={{ root: { width: 120 } }} />
                   <DefaultButton text="Thêm điểm" onClick={handleAddAnnotation} />
                </Stack>
                
                {/* Danh sách các điểm đã thêm */}
                {graphAnnotations.length > 0 && (
                  <div style={{ marginTop: "10px", display: "flex", gap: "10px", flexWrap: "wrap" }}>
                     {graphAnnotations.map((pt, idx) => (
                        <div key={idx} style={{ padding: "4px 8px", backgroundColor: "#e1dfdd", borderRadius: "4px", fontSize: "12px", display: "flex", alignItems: "center", gap: "6px" }}>
                           {/* Hiển thị Tên(X, Y) */}
                           {pt.name ? `${pt.name}` : ""}({Math.round(pt.x*100)/100}, {Math.round(pt.y*100)/100})
                           
                           <IconButton iconProps={{ iconName: 'Cancel' }} onClick={() => setGraphAnnotations(graphAnnotations.filter((_, i) => i !== idx))} styles={{ root: { width: 20, height: 20 }, icon: { fontSize: 10, color: "red" } }} />
                        </div>
                     ))}
                  </div>
                )}
              </div>

              <PrimaryButton text={t.insertToWord} iconProps={{ iconName: "LineChart" }} onClick={handleInsertGraph} style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }} />
            </div>
          )}

          {/* DIALOG QUẢN LÝ TỪ ĐIỂN */}
          <Dialog hidden={!isAddDictOpen} onDismiss={() => setIsAddDictOpen(false)} dialogContentProps={{ type: DialogType.normal, title: t.addDictTitle, subText: t.addDictSubText }}>
            <ChoiceGroup options={[{ key: 'inline', text: t.inlineOpt }, { key: 'macro', text: t.macroOpt }]} selectedKey={dictType} onChange={(_, option) => setDictType(option?.key as "inline" | "macro")} />
            <TextField label={t.shortcutLabel} placeholder={t.shortcutPlaceholder} value={dictKey} onChange={(_, val) => setDictKey(val || "")} styles={{ root: { marginTop: 10 }, fieldGroup: { height: 40 }, field: { fontSize: "15px", padding: "8px" } }} />
            <DialogFooter>
              <PrimaryButton onClick={handleSaveToDict} text={t.saveBtn} />
              <DefaultButton onClick={() => setIsAddDictOpen(false)} text={t.cancelBtn} />
            </DialogFooter>
          </Dialog>

          <Panel isOpen={isEditDictOpen} onDismiss={() => setIsEditDictOpen(false)} type={PanelType.medium} headerText={t.manageDictTitle} closeButtonAriaLabel="Close">
            <SearchBox placeholder={t.searchPlaceholder} value={searchTerm} onChange={(_, newValue) => setSearchTerm(newValue || "")} styles={{ root: { marginBottom: "16px", marginTop: "10px" } }} />
            <div style={{ maxHeight: "60vh", overflowY: "auto", paddingRight: "8px", overflowX: "hidden" }}>
              {["inline", "macro"].map((type) => {
                const filteredItems = editItems.filter(item => item.type === type && (item.key.toLowerCase().includes(searchTerm.toLowerCase()) || item.value.toLowerCase().includes(searchTerm.toLowerCase())));
                if (filteredItems.length === 0) return null;
                const displayItems = filteredItems.slice(0, 10);
                return (
                  <div key={type} style={{ marginBottom: "20px" }}>
                    <Label style={{ color: "#107c41", borderBottom: "1px solid #c8c6c4", paddingBottom: "4px", marginBottom: "8px" }}> {type === "inline" ? t.inlineHeader : t.macroHeader} </Label>
                    {displayItems.map((item) => (
                      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }} key={item.id} style={{ marginBottom: "12px" }}>
                        <TextField value={item.key} placeholder="Code..." onChange={(_, val) => handleItemChange(item.id, "key", val || "")} styles={{ root: { width: "100px" }, field: { fontWeight: "bold" } }} />
                        <TextField multiline autoAdjustHeight value={item.value} onChange={(_, val) => handleItemChange(item.id, "value", val || "")} styles={{ root: { flexGrow: 1 }, field: { fontFamily: "Consolas, monospace", fontSize: "13px" } }} />
                        <IconButton iconProps={{ iconName: 'Delete' }} title={t.deleteTitle} onClick={() => handleItemDelete(item.id)} styles={{ root: { color: "#d13438" }, rootHovered: { backgroundColor: "#fde7e9" } }} />
                      </Stack>
                    ))}
                    {filteredItems.length > 10 && (<Label style={{ fontStyle: "italic", color: "#d13438", fontSize: "12px", marginTop: "4px" }}> {t.showingResults.replace("{0}", filteredItems.length.toString())} </Label>)}
                  </div>
                );
              })}
              {editItems.length === 0 && (<Label style={{ fontStyle: "italic", color: "#a19f9d" }}>{t.emptyDict}</Label>)}
            </div>
            <Stack horizontal tokens={{ childrenGap: 10 }} style={{ marginTop: "20px" }}>
              <PrimaryButton iconProps={{ iconName: 'Save' }} text={t.saveChanges} onClick={handleSaveEditedDict} />
              <DefaultButton text={t.cancelBtn} onClick={() => setIsEditDictOpen(false)} />
            </Stack>
          </Panel>

        </Stack>
      </div>
    </ThemeProvider>
  );
}