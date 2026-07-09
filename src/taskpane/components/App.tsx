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


// ==========================================
// THÊM: Từ điển dịch UI
// ==========================================
type Language = "vi" | "en";
const uiText = {
  vi: {
    webWarning: "Phiên bản Web không được hỗ trợ chèn vào Word!\nĐể chèn công thức, vui lòng nhấn nút \"Mở trong ứng dụng trên máy tính\" (Open in Desktop App) trên thanh công cụ của Word.",
    mathLiveLabel: "Math Live:",
    latexLabel: "LaTeX:",
    importDict: "Nhập Từ điển",
    exportDict: "Xuất Từ điển",
    editDict: "Sửa Từ điển",
    saveToDict: "Lưu vào Từ điển",
    modeLabel: "Chế độ:",
    eEditorOnly: "E-Editor",
    eEditor: "E-Editor",
    office: "Office",
    officeMathWarning: "⚠️ Office Math không được hỗ trợ. Chỉ chèn kiểu E-Editor.",
    insertBtn: "Chèn / Cập nhật",
    computeLabel: "Tính toán:",
    evaluate: "Tính giá trị",
    simplify: "Rút gọn",
    expand: "Khai triển",
    expandAll: "Khai triển hết",
    factor: "Nhân tử",
    solve: "Tìm nghiệm",
    resultLabel: "Kết quả:",
    insertToWord: "Chèn vào Word",
    previewVertical: "Xem trước phép tính dọc:",
    previewVariation: "Xem trước Bảng biến thiên:",
    addDictTitle: "Thêm vào từ điển",
    addDictSubText: "Công thức này sẽ được lưu lại để bạn gõ nhanh lần sau.",
    inlineOpt: "Inline (Tự động chuyển đổi khi gõ)",
    macroOpt: "Macro (Gõ kèm dấu \\ phía trước để gợi ý)",
    shortcutLabel: "Phím tắt (không gõ '\\' ở trước)",
    shortcutPlaceholder: "VD: ptb, hpt...",
    saveBtn: "Lưu",
    cancelBtn: "Hủy",
    manageDictTitle: "Quản lý Từ điển",
    searchPlaceholder: "Tìm phím tắt hoặc mã LaTeX...",
    inlineHeader: "Phím tắt gõ nhanh (Inline)",
    macroHeader: "Lệnh Macro (\\)",
    deleteTitle: "Xóa phím tắt này",
    showingResults: "* Đang hiển thị 10 / {0} kết quả. Hãy nhập thêm từ khóa để tìm chính xác hơn.",
    emptyDict: "Từ điển đang trống.",
    saveChanges: "Lưu thay đổi",
    enterShortcutAlert: "Vui lòng nhập phím tắt!",
    importSuccessAlert: "Nhập từ điển thành công! Các phím tắt cũ của bạn vẫn được giữ nguyên.",
    importStructErrorAlert: "Cấu trúc file JSON không đúng chuẩn từ điển (Thiếu inline hoặc macro)!",
    importReadErrorAlert: "Lỗi đọc file JSON! Hãy kiểm tra lại định dạng file.",
    computeSystemNoSolution: "\\text{Hệ vô nghiệm hoặc số liệu không hợp lệ}",
    computeNotSupported: "\\text{Phiên bản thư viện chưa hỗ trợ hàm này}",
    computeSolveNotSupported: "\\text{Chưa hỗ trợ giải phương trình này hoặc vô nghiệm}",
    computeError: "Lỗi: Không thể tính toán biểu thức này.",
    derivative: "Đạo hàm",
    integrate: "Tích phân",
    limit: "Giới hạn"
  },
  en: {
    webWarning: "Web version does not support inserting into Word!\nTo insert equations, please click \"Open in Desktop App\" on the Word toolbar.",
    mathLiveLabel: "Math Live:",
    latexLabel: "LaTeX:",
    importDict: "Import Dict",
    exportDict: "Export Dict",
    editDict: "Edit Dict",
    saveToDict: "Save to Dict",
    modeLabel: "Mode:",
    eEditorOnly: "E-Editor (Only)",
    eEditor: "E-Editor",
    office: "Office",
    officeMathWarning: "⚠️ Office Math is not supported. Insert E-Editor type only.",
    insertBtn: "Insert / Update",
    computeLabel: "Compute:",
    evaluate: "Evaluate",
    simplify: "Simplify",
    expand: "Expand",
    expandAll: "Expand All",
    factor: "Factor",
    solve: "Solve",
    resultLabel: "Result:",
    insertToWord: "Insert into Word",
    previewVertical: "Preview Vertical Math:",
    previewVariation: "Preview Variation Table:",
    addDictTitle: "Add to Dictionary",
    addDictSubText: "This equation will be saved for quick typing next time.",
    inlineOpt: "Inline (Auto-convert while typing)",
    macroOpt: "Macro (Type with '\\' prefix for suggestions)",
    shortcutLabel: "Shortcut key (no '\\' prefix)",
    shortcutPlaceholder: "e.g., ptb, hpt...",
    saveBtn: "Save",
    cancelBtn: "Cancel",
    manageDictTitle: "Dictionary Manager",
    searchPlaceholder: "Search shortcut or LaTeX...",
    inlineHeader: "Inline Shortcuts",
    macroHeader: "Macro Commands (\\)",
    deleteTitle: "Delete this shortcut",
    showingResults: "* Showing 10 / {0} results. Type more keywords to refine search.",
    emptyDict: "Dictionary is empty.",
    saveChanges: "Save Changes",
    enterShortcutAlert: "Please enter a shortcut key!",
    importSuccessAlert: "Dictionary imported successfully! Your old shortcuts are kept.",
    importStructErrorAlert: "Invalid JSON structure for dictionary!",
    importReadErrorAlert: "Error reading JSON file! Please check the file format.",
    computeSystemNoSolution: "\\text{No solution or invalid data}",
    computeNotSupported: "\\text{Library version does not support this function}",
    computeSolveNotSupported: "\\text{Equation type not supported or no solution}",
    computeError: "Error: Cannot compute this expression.",
    derivative: "Derivative",
    integrate: "Integrate",
    limit: "Limit"
  }
};

export default function App() {
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const lastEqIdRef = useRef<string>("");
  const [currentLatex, setCurrentLatex] = useState("");

  const [insertMode, setInsertMode] = useState<string>(() => {
    return localStorage.getItem("math_insert_mode") || "mathml";
  });


  // --- BẮT ĐẦU: STATE VÀ LOGIC CHO COMPUTE ENGINE ---
  const [computeResult, setComputeResult] = useState<string>("");

  const defaultInlineRef = useRef<any>(null);
  const defaultMacrosRef = useRef<any>(null);

  const [statusMessage, setStatusMessage] = useState<{ text: string, type: MessageBarType } | null>(null);

  const [language, setLanguage] = useState<Language>(() => {
    // Đọc từ localStorage, nếu chưa có thì mặc định là 'vi'
    return (localStorage.getItem("math_user_lang") as Language) || "vi";
  });

  const t = uiText[language]; // Biến 't' này sẽ dùng để gọi chữ ra UI

  const handleCompute = (action: "evaluate" | "simplify" | "expand" | "expandAll" | "factor" | "solve" |"derivative" |"integrate"| "limit") => {
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
        case "derivative": // Đạo hàm (vd: d/dx(x^2) = 2x)
          resultLatex = ce.expr(["Derivative", expr]).evaluate().latex;
          break;
        case "integrate": // Tích phân / Nguyên hàm
          // 1. Tự động tìm ẩn số để truyền vào engine (VD: x, t, y)
          let intVar = "x";
          if (expr.unknowns && expr.unknowns.length > 0) {
            intVar = expr.unknowns[0];
          } else if (currentLatex.includes("t")) {
            intVar = "t";
          } else if (currentLatex.includes("y")) {
            intVar = "y";
          }

          // 2. Chặn lỗi bấm nhầm: Nếu người dùng đã gõ sẵn ký hiệu \int 
          // (ví dụ: \int x^2 dx) nhưng lại bấm nhầm nút "Integrate" thay vì "Evaluate"
          if (currentLatex.includes("\\int")) {
            const evalInt = expr.evaluate();
            if (evalInt && ((evalInt as any).head === "Integrate" || evalInt.latex.includes("\\int"))) {
              resultLatex = t.computeNotSupported;
            } else {
              resultLatex = evalInt.latex;
            }
          } 
          // 3. Xử lý chuẩn: Bôi đen biểu thức (vd: x^2) và bấm nút Integrate
          else {
            const evalInt = ce.expr(["Integrate", expr, intVar]).evaluate();
            
            // Kiểm tra xem CortexJS có giải được không
            if (evalInt && ((evalInt as any).head === "Integrate" || evalInt.latex.includes("\\int"))) {
              resultLatex = t.computeNotSupported;
            } else {
              // Bổ sung hằng số tích phân C cho chuẩn chỉnh về mặt toán học
              resultLatex = evalInt.latex + " + C";
            }
          }
          break;
        case "limit": // Giới hạn (Hỗ trợ vô cực và 1 bên)
          const limitRegex = /\\lim\s*_\s*\{\s*([a-zA-Z])\s*(?:\\to|\\rightarrow)\s*([^}]+)\s*\}(.*)/;
          const limitMatch = currentLatex.match(limitRegex);

          if (limitMatch) {
            const varName = limitMatch[1];
            let targetLatex = limitMatch[2].trim();
            const funcLatex = limitMatch[3];

            let direction = 0;
            if (targetLatex.includes("^+") || targetLatex.includes("^{+}")) {
              direction = 1;
              targetLatex = targetLatex.replace(/\^\{\+\}|\^\+/g, "");
            } else if (targetLatex.includes("^-") || targetLatex.includes("^{-}")) {
              direction = -1;
              targetLatex = targetLatex.replace(/\^\{-\}|\^\-/g, "");
            }

            const funcExpr = ce.parse(funcLatex);
            const targetExpr = ce.parse(targetLatex);

            try {
              // CHIẾN LƯỢC MỚI: ƯU TIÊN THẾ SỐ TRƯỚC (NUMERIC FIRST)
              let testValueStr = targetLatex;

              if (targetLatex.includes("\\infty")) {
                if (targetLatex.includes("-")) {
                  testValueStr = "-10000000"; // Thế -10 triệu
                } else {
                  testValueStr = "10000000";  // Thế 10 triệu
                }
              } else {
                let numTarget = parseFloat((targetExpr as any).numericValue || targetExpr.value || targetLatex);
                if (!isNaN(numTarget)) {
                  if (direction === 1) numTarget += 1e-7;
                  else if (direction === -1) numTarget -= 1e-7;
                  else numTarget += 1e-8; // Lệch siêu nhỏ để tránh lỗi chia 0
                  testValueStr = numTarget.toString();
                }
              }

              // Ép thư viện tính ra số thập phân thực tế (Dùng .N())
              const subExpr = funcExpr.subs({ [varName]: ce.parse(testValueStr) });
              const numericValObj = subExpr.N();
              
              let numericRes = NaN;
              if (numericValObj && typeof (numericValObj as any).numericValue === 'number') {
                numericRes = (numericValObj as any).numericValue;
              } else {
                numericRes = parseFloat(numericValObj.toString());
              }

              // Phân tích kết quả số học
              if (!isNaN(numericRes)) {
                if (numericRes > 1e5 || !isFinite(numericRes)) {
                  resultLatex = "+\\infty";
                } else if (numericRes < -1e5) {
                  resultLatex = "-\\infty";
                } else {
                  // Làm tròn khử sai số (vd: 1e-7 sẽ tự động thành 0)
                  resultLatex = (Math.round(numericRes * 10000) / 10000).toString();
                }
              } else {
                // Dự phòng cuối: Nếu thế số thất bại, mới gọi Symbolic Limit của CortexJS
                let evalLimit = direction !== 0 
                  ? ce.expr(["Limit", funcExpr, varName, targetExpr, direction]).evaluate()
                  : ce.expr(["Limit", funcExpr, varName, targetExpr]).evaluate();
                
                if (evalLimit && ((evalLimit as any).head === "Limit" || evalLimit.latex.includes("\\lim"))) {
                  resultLatex = t.computeNotSupported;
                } else {
                  resultLatex = evalLimit.latex;
                }
              }
            } catch (e) {
              resultLatex = t.computeNotSupported;
            }
          } else {
            // Dành cho biểu thức không chứa cấu trúc \lim
            const evalResult = expr.evaluate();
            if (evalResult && ((evalResult as any).head === "Limit" || evalResult.latex.includes("\\lim"))) {
               resultLatex = t.computeNotSupported;
            } else {
               resultLatex = evalResult ? evalResult.latex : t.computeNotSupported;
            }
          }
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
                resultLatex = t.computeSystemNoSolution; // "Hệ vô nghiệm hoặc số liệu không hợp lệ"
              }
            } else {
              resultLatex = t.computeNotSupported; // "Phiên bản thư viện chưa hỗ trợ hàm này"
            }
          } else {
            // ---> GIẢI PHƯƠNG TRÌNH 1 ẨN (Giữ nguyên code cũ đang chạy tốt) <---
            const v = vars[0];
            const solveResult = ce.expr(["Solve", expr, v]).evaluate();

            if (solveResult && solveResult.latex.includes("Solve")) {
              resultLatex = t.computeSolveNotSupported;
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
      setComputeResult("t.computeError"); //Lỗi không thể tính toán biểu thức này
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
      localStorage.setItem("math_insert_mode", "svg"); // LƯU VÀO BỘ NHỚ
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

    //Bỏ qua tính dọc nếu là hàm số
    if (bieuThuc.includes('y=') || bieuThuc.includes('f(x)=')) return null;

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
        svgString += `<text x="${rightEdge}" y="${currentY + 8}" text-anchor="end" ${FONT_STYLE} fill="black">${ketQuaCuoi}</text>`;

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

  // ==========================================
  // TÍNH NĂNG MỚI: PREVIEW BẢNG BIẾN THIÊN (DẠNG SVG CHUẨN, NGHIỆM CHÍNH XÁC)
  // ==========================================
  // ==========================================
  // TÍNH NĂNG MỚI: PREVIEW BẢNG BIẾN THIÊN (HỖ TRỢ HÀM PHÂN THỨC, TIỆM CẬN ĐỨNG)
  // ==========================================
  const variationTablePreview = useMemo(() => {
    if (!currentLatex) return null;
    const latex = currentLatex.trim();
    if (!latex.startsWith("y=") && !latex.startsWith("f(x)=")) return null;

    try {
      const rhs = latex.substring(latex.indexOf("=") + 1).trim();
      const expr = ce.parse(rhs);

      // 1. HÀM QUÉT NGHIỆM ĐỆ QUY (Tìm sâu các giá trị số trong MathJSON)
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

      // 2. TÌM ĐIỂM GIÁN ĐOẠN (MẪU SỐ BẰNG 0) DÀNH CHO HÀM PHÂN THỨC
      let rawDiscontinuities: number[] = [];
      const fracRegex = /\\frac\{([^}]+)\}\{([^}]+)\}/g;
      let match;
      while ((match = fracRegex.exec(rhs)) !== null) {
          const denomStr = match[2];
          const denomExpr = ce.parse(denomStr);
          // Tìm nghiệm của mẫu số
          const solveDenom = ce.expr(["Solve", ["Equal", denomExpr, 0], "x"]).evaluate();
          extractRootsRecursively(solveDenom, rawDiscontinuities);
      }

      // 3. TÍNH ĐẠO HÀM Y' VÀ TÌM CỰC TRỊ Y' = 0
      const deriv = ce.expr(["Derivative", expr, "x"]).evaluate().simplify();
      const solveRes = ce.expr(["Solve", ["Equal", deriv, 0], "x"]).evaluate();
      
      let rawRoots: number[] = [];
      extractRootsRecursively(solveRes, rawRoots);

      // 4. GỘP NGHIỆM VÀ ĐIỂM GIÁN ĐOẠN, LỌC TRÙNG VÀ SẮP XẾP TĂNG DẦN
      let allPoints = [...rawRoots, ...rawDiscontinuities].sort((a, b) => a - b);
      let uniquePoints: { val: number, isDis: boolean }[] = [];
      for(let i=0; i<allPoints.length; i++) {
          // Lọc các điểm sát nhau (khử sai số làm tròn)
          if (i > 0 && Math.abs(allPoints[i] - allPoints[i-1]) < 1e-6) continue;
          const val = allPoints[i];
          // Kiểm tra xem điểm này có nằm trong danh sách làm mẫu số bằng 0 không
          const isDis = rawDiscontinuities.some(d => Math.abs(d - val) < 1e-6);
          uniquePoints.push({ val, isDis });
      }

      // Hàm ép tính toán kết quả ra số thực
      const getNum = (exprNode: any) => {
          const val = exprNode.evaluate();
          if (val.numericValue !== undefined && val.numericValue !== null) return val.numericValue;
          return parseFloat(val.toString());
      };

      // 5. XÉT DẤU Y' TỪNG KHOẢNG (Bằng cách lấy điểm thử Test Point)
      const signOfDerivAt = (xVal: number) => {
         const res = getNum(deriv.subs({ x: xVal }));
         if (isNaN(res) || !isFinite(res)) return "+"; // Fallback phòng hờ lỗi chia 0
         return res >= 0 ? "+" : "-";
      };

      let signs: string[] = [];
      for (let i = 0; i <= uniquePoints.length; i++) {
         let testPoint = 0;
         if (i === 0) testPoint = uniquePoints.length > 0 ? uniquePoints[0].val - 1 : 0;
         else if (i === uniquePoints.length) testPoint = uniquePoints[i-1].val + 1;
         else testPoint = (uniquePoints[i-1].val + uniquePoints[i].val) / 2;
         signs.push(signOfDerivAt(testPoint)); 
      }

      // 6. TÍNH GIỚI HẠN / GIÁ TRỊ TẠI CÁC ĐIỂM (Cực trị & Tiệm cận đứng)
      const getLimit = (xVal: number, direction = 0) => {
         let testX = xVal;
         if (xVal === -1000) testX = -1e6;
         else if (xVal === 1000) testX = 1e6;
         else if (direction === -1) testX = xVal - 1e-6; // Tiến về bên trái
         else if (direction === 1) testX = xVal + 1e-6;  // Tiến về bên phải

         const res = getNum(expr.subs({ x: testX }));
         if (isNaN(res)) return "?";
         if (res > 1e4 || !isFinite(res)) return (res > 0 || Object.is(res, Infinity)) ? "+\\infty" : "-\\infty";
         if (res < -1e4) return "-\\infty";
         return Math.round(res * 100) / 100;
      };

      // 7. XÂY DỰNG CẤU TRÚC CỘT (COLUMNS DATA)
      let cols: any[] = [];
      //cols.push({ type: 'val', x: "-\\infty", isDis: false, y: getLimit(-1000) });
      cols.push({ type: 'val', x: "-\\infty", isDis: false, yPrime: "", y: getLimit(-1000) });
      
      for (let i = 0; i < uniquePoints.length; i++) {
         let pt = uniquePoints[i];
         cols.push({ type: 'interval', sign: signs[i] });

         if (pt.isDis) {
             // Điểm gián đoạn (Có giới hạn trái và giới hạn phải)
             cols.push({
                 type: 'val', x: pt.val, isDis: true,
                 yPrime: "", // Đã sửa: Xóa yPrime tại tiệm cận đứng
                 yLeft: getLimit(pt.val, -1),
                 yRight: getLimit(pt.val, 1)
             });
         } else {
             // Điểm cực đại / cực tiểu bình thường
             cols.push({
                 type: 'val', x: pt.val, isDis: false,
                 yPrime: "0", y: getLimit(pt.val, 0)
             });
         }
      }
      
      cols.push({ type: 'interval', sign: signs[signs.length - 1] });
      cols.push({ type: 'val', x: "+\\infty", isDis: false, y: getLimit(1000) });

      // 8. XÁC ĐỊNH TỌA ĐỘ Y (TOP/BOTTOM) ĐỂ VẼ CHỮ & MŨI TÊN KHÔNG BỊ CHỒNG LÊN NHAU
      let valCols = cols.filter(c => c.type === 'val');
      for(let i=0; i<valCols.length; i++) {
          let c = valCols[i];
          if (i === 0) {
              c.pos = signs[0] === '+' ? 'BOTTOM' : 'TOP';
          } else if (i === valCols.length - 1) {
              c.pos = signs[signs.length - 1] === '+' ? 'TOP' : 'BOTTOM';
          } else {
              let inSign = signs[i-1];
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

      // 9. VẼ BẢNG BIẾN THIÊN BẰNG SVG THUẦN TÚY
      const X_OFFSET = 50; 
      const COL_W = 60; 
      const W = X_OFFSET + cols.length * COL_W; 
      const H = 180; 

      let currentX = X_OFFSET;
      cols.forEach(c => { c.cx = currentX + COL_W / 2; currentX += COL_W; });

      let svgString = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="background-color: white;">`;
      
      // Đầu mũi tên đồ thị
      svgString += `<defs><marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="black" /></marker></defs>`;
      
      // Khung viền và các trục cắt ngang/dọc
      svgString += `<rect x="0" y="0" width="${W}" height="${H}" fill="none" stroke="black" stroke-width="1.5"/>`;
      svgString += `<line x1="0" y1="40" x2="${W}" y2="40" stroke="black" stroke-width="1.5"/>`;
      svgString += `<line x1="0" y1="80" x2="${W}" y2="80" stroke="black" stroke-width="1.5"/>`;
      svgString += `<line x1="${X_OFFSET}" y1="0" x2="${X_OFFSET}" y2="${H}" stroke="black" stroke-width="1.5"/>`;

      const FONT = `font-family="Times New Roman, serif" font-size="18px" text-anchor="middle" dominant-baseline="middle"`;
      
      // Nhãn các trục
      svgString += `<text x="${X_OFFSET/2}" y="20" ${FONT} font-style="italic">x</text>`;
      svgString += `<text x="${X_OFFSET/2}" y="60" ${FONT} font-style="italic">y'</text>`;
      svgString += `<text x="${X_OFFSET/2}" y="130" ${FONT} font-style="italic">y</text>`;

      // Render nội dung các cột
      cols.forEach((c, i) => {
          if (c.type === 'val') {
              // --- Xử lý dòng x ---
              let txtX = c.x === "-\\infty" ? "-\u221E" : (c.x === "+\\infty" ? "+\u221E" : c.x);
              svgString += `<text x="${c.cx}" y="20" ${FONT}>${txtX}</text>`;
              
              if (c.isDis) {
                  // VẼ TIỆM CẬN ĐỨNG (2 gạch dọc || song song)
                  svgString += `<line x1="${c.cx - 3}" y1="40" x2="${c.cx - 3}" y2="${H}" stroke="black" stroke-width="1.5"/>`;
                  svgString += `<line x1="${c.cx + 3}" y1="40" x2="${c.cx + 3}" y2="${H}" stroke="black" stroke-width="1.5"/>`;

                  // Hàng y (Giới hạn trái và phải ốp sát 2 bên đường tiệm cận)
                  let txtYL = c.yLeft === "-\\infty" ? "-\u221E" : (c.yLeft === "+\\infty" ? "+\u221E" : c.yLeft);
                  let txtYR = c.yRight === "-\\infty" ? "-\u221E" : (c.yRight === "+\\infty" ? "+\u221E" : c.yRight);
                  
                  let yDLeft = c.posLeft === 'TOP' ? 100 : 165;
                  let yDRight = c.posRight === 'TOP' ? 100 : 165;

                  svgString += `<text x="${c.cx - 20}" y="${yDLeft}" ${FONT} font-weight="bold">${txtYL}</text>`;
                  svgString += `<text x="${c.cx + 20}" y="${yDRight}" ${FONT} font-weight="bold">${txtYR}</text>`;
              } else {
                  // --- Xử lý hàm liên tục ---
                  //if (c.yPrime !== "") svgString += `<text x="${c.cx}" y="60" ${FONT}>${c.yPrime}</text>`;
                  if (c.yPrime) svgString += `<text x="${c.cx}" y="60" ${FONT}>${c.yPrime}</text>`;
                  
                  let txtY = c.y === "-\\infty" ? "-\u221E" : (c.y === "+\\infty" ? "+\u221E" : c.y);
                  let yD = c.pos === 'TOP' ? 100 : (c.pos === 'BOTTOM' ? 165 : 130);
                  svgString += `<text x="${c.cx}" y="${yD}" ${FONT} font-weight="bold">${txtY}</text>`;
              }
          } else {
              // --- Vẽ Dấu (+/-) & Mũi tên Đồ thị ---
              svgString += `<text x="${c.cx}" y="60" ${FONT} fill="red" font-weight="bold">${c.sign}</text>`;
              
              const prevVal = cols[i - 1];
              const nextVal = cols[i + 1];
              
              // Căn chỉnh tọa độ y để mũi tên nghiêng chuẩn
              let y1 = prevVal.isDis ? (prevVal.posRight === 'TOP' ? 112 : 152) : (prevVal.pos === 'TOP' ? 112 : (prevVal.pos === 'BOTTOM' ? 152 : 130));
              let y2 = nextVal.isDis ? (nextVal.posLeft === 'TOP' ? 112 : 152) : (nextVal.pos === 'TOP' ? 112 : (nextVal.pos === 'BOTTOM' ? 152 : 130));
              
              // Lùi điểm xuất phát/kết thúc mũi tên ra xa hơn 1 chút nếu gặp đường 2 gạch (||)
              let x1 = prevVal.cx + (prevVal.isDis ? 32 : 20);
              let x2 = nextVal.cx - (nextVal.isDis ? 38 : 24);
              
              svgString += `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="black" stroke-width="1.5" marker-end="url(#arrow)" />`;
          }
      });

      svgString += `</svg>`;

      return { svgString, mathWidthEx: W/14, mathHeightEx: H/14, bieuThuc: latex };
    } catch (err) {
      console.error("Lỗi vẽ BBT:", err);
      return null;
    }
  }, [currentLatex]);


  // HÀM CHÈN SVG BẢNG BIẾN THIÊN VÀO WORD
  const handleInsertVariationTable = async () => {
    if (!variationTablePreview) return;
    const { svgString, mathWidthEx, mathHeightEx, bieuThuc } = variationTablePreview;

    try {
      // Mã hóa SVG để Word đọc được
      let utf8Encoded = encodeURIComponent(svgString).replace(/%([0-9A-F]{2})/g, (_match, p1) => String.fromCharCode(parseInt(p1, 16)));
      let insertContent = window.btoa(utf8Encoded);

      const eqId = "EQT" + new Date().getTime();
      Office.context.document.settings.set(eqId, bieuThuc);
      await Office.context.document.settings.saveAsync();

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.load("font/size");
        await context.sync();

        const currentFontSize = range.font.size || 12;
        const exToPt = currentFontSize * 0.44;

        // Chèn ảnh dạng SVG
        const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);
        picture.width = mathWidthEx * exToPt;
        picture.height = mathHeightEx * exToPt;
        picture.lockAspectRatio = true;
        picture.altTextTitle = "Variation Table";
        picture.altTextDescription = `Math: ${bieuThuc} #[ID:${eqId}|SVG]`;

        const newCc = picture.getRange().insertContentControl();
        newCc.title = "LSTS_Math";
        newCc.tag = eqId + "|SVG";
        newCc.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();
      });
    } catch (err) {
      console.error("Lỗi chèn BBT: ", err);
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

      /*
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
      */



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


  // ==========================================
  // THAY ĐỔI NGÔN NGỮ HOẶC DICT: XỬ LÝ MACRO
  // ==========================================

  useEffect(() => {
    if (!mathfieldRef.current) return;

    const mf = mathfieldRef.current;

    // 1. Luôn bắt đầu từ bản gốc (để đảm bảo \frac, \sqrt... luôn tồn tại)
    // MathLive hỗ trợ gán lại object mới để đè lên các thay đổi cũ
    const baseInlines = { ...defaultInlineRef.current };
    const baseMacros = { ...defaultMacrosRef.current };

    // 2. Gom phím tắt người dùng tự thêm (Luôn luôn có)
    const userInlines = { ...userDict.inline };
    const userMacros = { ...upgradeMacros(userDict.macro) };

    if (language === "vi") {
      // TIẾNG VIỆT: Gộp bản gốc + Phím tắt người dùng + Phím tắt custom code
      mf.inlineShortcuts = {
        ...baseInlines,
        ...customInlineShortcuts,
        ...userInlines
      };
      mf.macros = {
        ...baseMacros,
        ...upgradeMacros(customMacros),
        ...userMacros
      };
    } else {
      // TIẾNG ANH: Gộp bản gốc + Phím tắt người dùng (KHÔNG gộp custom code)
      mf.inlineShortcuts = {
        ...baseInlines,
        ...userInlines
      };
      mf.macros = {
        ...baseMacros,
        ...userMacros
      };
    }
  }, [language, userDict]);


  const insertSymbol = (latexCode: string) => {
    if (mathfieldRef.current) {
      mathfieldRef.current.insert(latexCode);
      mathfieldRef.current.focus();
    }
  };

  const handleSaveToDict = () => {
    if (!dictKey.trim()) {
      //alert(t.enterShortcutAlert); //Vui lòng nhập phím tắt
      setStatusMessage({ text: t.enterShortcutAlert, type: MessageBarType.warning });
      return;
    }
    const updatedDict = { ...userDict };
    updatedDict[dictType][dictKey.trim()] = currentLatex;
    setUserDict(updatedDict);
    localStorage.setItem("math_user_dict", JSON.stringify(updatedDict));

    /*
    if (mathfieldRef.current) {
      if (dictType === "inline") {
        mathfieldRef.current.inlineShortcuts = { ...mathfieldRef.current.inlineShortcuts, ...updatedDict.inline };
      } else {
        mathfieldRef.current.macros = { ...mathfieldRef.current.macros, ...upgradeMacros(updatedDict.macro) };
      }
    }
      */
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
          /*
          if (mathfieldRef.current) {
            mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...mergedDict.inline };
            mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(mergedDict.macro) };
          }
            */
          setTimeout(() => {
            //alert(t.importSuccessAlert);
            setStatusMessage({ text: t.importSuccessAlert, type: MessageBarType.success });
          }, 100);
        } else {
          //alert(t.importStructErrorAlert);
          setStatusMessage({ text: t.importStructErrorAlert, type: MessageBarType.error });
        }
      } catch (error) {
        //alert(t.importReadErrorAlert);
        setStatusMessage({ text: t.importReadErrorAlert, type: MessageBarType.error });
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
    /*
    if (mathfieldRef.current) {
      mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...newDict.inline };
      mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(newDict.macro) };
    }
      */
    setIsEditDictOpen(false);
  };

  return (
    <ThemeProvider theme={officeGreenTheme}>
      <div style={{ padding: "8px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8", overflowY: "auto" }}>
        <div style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: "8px"
        }}>
          <Label style={{ fontSize: "14px", fontWeight: "bold", margin: 0, color: "#107c41" }}>
            Easy Efficient Equation Editor
          </Label>

          <Toggle
            onText="Tiếng Việt"
            offText="English"
            checked={language === 'vi'}
            onChange={(_, checked) => {
              const newLang = checked ? 'vi' : 'en';
              setLanguage(newLang);
              localStorage.setItem("math_user_lang", newLang);
            }}
            styles={{ root: { margin: 0 } }} // Bỏ margin để căn chỉnh chính xác
          />
        </div>
        <Stack tokens={{ childrenGap: 12 }}>
          {isOnline && (
            <MessageBar messageBarType={MessageBarType.blocked} isMultiline={true}>
              <span style={{ whiteSpace: "pre-line" }}>{t.webWarning}</span>
            </MessageBar>
          )}
          {statusMessage && (
            <MessageBar
              messageBarType={statusMessage.type}
              onDismiss={() => setStatusMessage(null)}
              isMultiline={false}
            >
              {statusMessage.text}
            </MessageBar>
          )}

          <div style={{ height: "145px" }}>
            <MathPalette onInsert={insertSymbol} language={language} />
          </div>

          <Label style={{ fontSize: "14px", fontWeight: "600" }}>{t.mathLiveLabel}</Label>
          <div ref={containerRef} id="mathfield-container" />

          <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: "4px" }}>
              <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>{t.latexLabel}</Label>

              <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                <input type="file" accept=".json" ref={fileInputRef} style={{ display: "none" }} onChange={handleImportDict} />

                <DefaultButton
                  iconProps={{ iconName: 'Upload' }}
                  title={t.importDict}
                  onClick={() => fileInputRef.current?.click()}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Download' }}
                  title={t.exportDict}
                  onClick={handleExportDict}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Edit' }}
                  title={t.editDict}
                  onClick={handleOpenEditDict}
                  styles={{ root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, icon: { fontSize: 12 } }}
                />
                <DefaultButton
                  iconProps={{ iconName: 'Save' }}
                  text="Save to Dict"
                  title={t.saveToDict}
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
                label={t.modeLabel}
                inlineLabel
                onText={isUnsupported ? t.eEditorOnly : t.eEditor}
                offText={t.office}
                checked={insertMode === "svg"}
                onChange={(_ev, checked) => {
                  const newMode = checked ? "svg" : "mathml";
                  setInsertMode(newMode);
                  localStorage.setItem("math_insert_mode", newMode); // LƯU VÀO BỘ NHỚ
                }}
                disabled={isUnsupported}
                styles={{ root: { marginTop: "10px", padding: "10px 0px", marginBottom: isUnsupported ? "4px" : "10px" } }}
              />

              {isUnsupported && (
                <Label style={{ fontSize: "12px", color: "#d13438", marginTop: "0px", marginBottom: "10px", lineHeight: "1.2" }}>
                  {t.officeMathWarning}
                </Label>
              )}
            </Stack>

            <DefaultButton
              primary
              text={t.insertBtn}
              onClick={handleInsertEquation}
              style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", fontWeight: "bold", borderRadius: "8px" }}
              title="Insert new equation or update selected equation"
            />
          </Stack>

          {/* ========================================== */}
          {/* KHU VỰC TÍNH TOÁN BẰNG COMPUTE ENGINE      */}
          {/* ========================================== */}
          <div style={{ marginTop: "12px", padding: "12px", backgroundColor: "#f3f2f1", borderRadius: "6px", border: "1px solid #e1dfdd" }}>
            <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", padding: 0 }}>{t.computeLabel}</Label>

            <Stack horizontal wrap tokens={{ childrenGap: 8 }} style={{ marginBottom: "0px", marginTop: "8px" }}>
              <DefaultButton iconProps={{ iconName: 'Calculator' }} text={t.evaluate} onClick={() => handleCompute("evaluate")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Broom' }} text={t.simplify} onClick={() => handleCompute("simplify")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'ExploreContent' }} text={t.expand} onClick={() => handleCompute("expand")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'MiniExpand' }} text={t.expandAll} onClick={() => handleCompute("expandAll")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'Merge' }} text={t.factor} onClick={() => handleCompute("factor")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton
                iconProps={{ iconName: 'Variable' }}
                text={t.solve}
                onClick={() => handleCompute("solve")}
                style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }}
              />
              <DefaultButton iconProps={{ iconName: 'CalculatorAddition' }} text={t.derivative} onClick={() => handleCompute("derivative")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'CalculatorSubtract' }} text={t.integrate} onClick={() => handleCompute("integrate")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
              <DefaultButton iconProps={{ iconName: 'CalculatorMultiply' }} text={t.limit} onClick={() => handleCompute("limit")} style={{ padding: "0px 5px", minWidth: "100px", borderRadius: "6px" }} />
            </Stack>

            {computeResult && (
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: "16px", borderTop: "1px dashed #c8c6c4", paddingTop: "12px" }}>
                <Label style={{ color: "#605e5c", margin: 0 }}>{t.resultLabel}</Label>

                {/* Dùng thẻ math-field chỉ đọc để render công thức kết quả đẹp mắt */}
                <div style={{ padding: "10px", backgroundColor: "white", borderRadius: "4px", fontSize: "22px", minHeight: "50px" }}>
                  <math-field read-only style={{ border: "none", outline: "none", boxShadow: "none" }}>
                    {computeResult}
                  </math-field>
                </div>

                <PrimaryButton
                  text={t.insertToWord}
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
                {t.previewVertical}
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
                text={t.insertToWord}
                iconProps={{ iconName: "CalculatorAddition" }}
                onClick={handleInsertVerticalMath}
                style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }}
              />
            </div>
          )}

          {/* DÁN KHỐI GIAO DIỆN BẢNG BIẾN THIÊN VÀO ĐÂY */}
          {variationTablePreview && (
            <div style={{
              marginTop: "4px", padding: "16px", backgroundColor: "#f3f2f1", borderRadius: "6px",
              border: "1px solid #e1dfdd", display: "flex", flexDirection: "column", alignItems: "center"
            }}>
              <Label style={{ fontWeight: "bold", color: "#107c41", fontSize: "14px", alignSelf: "flex-start" }}>
                {t.previewVariation}
              </Label>

              {/* Đã đổi từ htmlString sang svgString ở dòng dưới đây */}
              <div style={{
                backgroundColor: "white", padding: "10px", borderRadius: "8px", width: "100%",
                boxShadow: "0 2px 8px rgba(0,0,0,0.1)", margin: "12px 0", overflowX: "auto"
              }}
                dangerouslySetInnerHTML={{ __html: variationTablePreview.svgString }}
              />

              <PrimaryButton
                text={t.insertToWord}
                iconProps={{ iconName: "Table" }}
                onClick={handleInsertVariationTable}
                style={{ height: "40px", borderRadius: "6px", backgroundColor: "#0078d4", borderColor: "transparent", width: "100%", maxWidth: "300px" }}
              />
            </div>
          )}

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