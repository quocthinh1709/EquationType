import * as React from "react";
import { useEffect, useRef, useState, useCallback } from "react";
//import { DefaultButton, Stack, Label } from "@fluentui/react";
import { DefaultButton,IconButton, Stack, Label, ChoiceGroup, IChoiceGroupOption, Toggle, MessageBar, MessageBarType } from "@fluentui/react";
//import { DefaultButton, IconButton, Stack, Label, ChoiceGroup, IChoiceGroupOption, Toggle, Dialog, DialogType, DialogFooter, PrimaryButton, TextField } from "@fluentui/react";
import { Dialog, DialogType, DialogFooter, PrimaryButton, TextField, Panel, PanelType, SearchBox, ThemeProvider, createTheme } from "@fluentui/react";
import { initializeIcons } from '@fluentui/react/lib/Icons';

import { MathfieldElement } from "mathlive";
import "mathlive";
import MathPalette from "./MathPalette";
import { customInlineShortcuts, customMacros } from "./mathShortcuts";
import { unsupportedMathMlCommands } from "./unsupportedCommands";

// Khởi tạo bộ Icon của Fluent UI
initializeIcons();
// Định nghĩa bộ Theme Xanh lá chuẩn của Office
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

// Đảm bảo load Font nội bộ chuẩn xác
MathfieldElement.fontsDirectory = "./fonts";



export default function App() {
  // BƯỚC 1: KIỂM TRA MÔI TRƯỜNG LÀ WORD ONLINE
  const isOnline = Office.context.platform === Office.PlatformType.OfficeOnline;

  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const [currentLatex, setCurrentLatex] = useState("");

  // THÊM BIẾN TRẠNG THÁI LƯU CHẾ ĐỘ CHÈN
  const [insertMode, setInsertMode] = useState<string>("mathml");


  // --- STATE TỪ ĐIỂN ---
  // Khởi tạo từ điển từ localStorage (nếu có)
  const [userDict, setUserDict] = useState<{ inline: Record<string, string>, macro: Record<string, string> }>(() => {
    try {
      const saved = localStorage.getItem("math_user_dict");
      return saved ? JSON.parse(saved) : { inline: {}, macro: {} };
    } catch (e) {
      return { inline: {}, macro: {} };
    }
  });

  // Quản lý hiển thị Popup Thêm từ điển
  const [isAddDictOpen, setIsAddDictOpen] = useState(false);
  const [dictKey, setDictKey] = useState("");
  const [dictType, setDictType] = useState<"inline" | "macro">("inline");
  const fileInputRef = useRef<HTMLInputElement>(null);
  // ---------------------
  // Định nghĩa cấu trúc của 1 item khi đang sửa
  interface EditItem { id: string; type: "inline" | "macro"; key: string; value: string; }
  // Quản lý hiển thị Panel sửa từ điển
  // State quản lý danh sách sửa
  const [isEditDictOpen, setIsEditDictOpen] = useState(false);
  const [editItems, setEditItems] = useState<EditItem[]>([]);
  const [searchTerm, setSearchTerm] = useState("");

  // Kiểm tra xem currentLatex có chứa bất kỳ chuỗi nào trong danh sách cấm hay không
  const isUnsupported = unsupportedMathMlCommands.some(cmd => currentLatex.includes(cmd));
  // Tự động gạt công tắc sang "Ảnh" nếu phát hiện lệnh cấm
  useEffect(() => {
    if (isUnsupported && insertMode === "mathml") {
      setInsertMode("svg");
    }
  }, [isUnsupported, insertMode]);

  // ĐỊNH NGHĨA 2 LỰA CHỌN CHO NGƯỜI DÙNG
  const modeOptions: IChoiceGroupOption[] = [
    { key: "mathml", text: "Mã MathML" },
    { key: "svg", text: "Ảnh" }
  ];

  // ==========================================
  // BỘ LỌC TỰ ĐỘNG: Biến mọi chuỗi Macro tĩnh thành Macro Động (Có gợi ý + Cho phép Sửa)
  // ==========================================
  const upgradeMacros = (macroDict: Record<string, any>) => {
    const upgraded: Record<string, any> = {};
    Object.keys(macroDict).forEach(key => {
      const val = macroDict[key];
      if (typeof val === "string") {
        upgraded[key] = {
          def: val,                 // Lấy mã LaTeX gốc
          captureSelection: false,  // Mở khóa click chuột
          expand: true              // Ép bung mã, tự động nhận diện #0 và #?
        };
      } else {
        upgraded[key] = val;
      }
    });
    return upgraded;
  };

  // ==========================================
  // HÀM 1: ĐỌC NGƯỢC TỪ WORD VÀO ADD-IN (KHI BÔI ĐEN/CLICK)
  // ==========================================
  const handleReadEquation = useCallback(async () => {
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();

        // Bắt trường hợp nhấp nháy chuột bên trong
        const cc = range.parentContentControlOrNullObject;
        cc.load("title, tag");

        // Bắt trường hợp bôi đen toàn bộ công thức
        const ccs = range.contentControls;
        ccs.load("items/title, items/tag");

        await context.sync();

        let eqId = "";

        // Ưu tiên đọc khi bôi đen
        if (ccs.items.length > 0 && ccs.items[0].title === "LSTS_Math") {
          eqId = ccs.items[0].tag;
        }
        // Sau đó đọc khi click chuột
        else if (!cc.isNullObject && cc.title === "LSTS_Math") {
          eqId = cc.tag;
        }

        // Nếu tìm thấy ID, lôi mã LaTeX từ bộ nhớ Word ra
        if (eqId) {
          const latexCode = Office.context.document.settings.get(eqId);
          if (latexCode && mathfieldRef.current) {
            // Cập nhật lại khung hình và khung code
            if (mathfieldRef.current.value !== latexCode) {
              mathfieldRef.current.value = latexCode;
              setCurrentLatex(latexCode);
            }
          }
        }
      });
    } catch (error) {
      // Bỏ qua lỗi ngầm để trải nghiệm mượt mà
    }
  }, []);

  // ==========================================
  // HÀM 2: CHÈN HOẶC GHI ĐÈ CÔNG THỨC VÀO WORD
  // ==========================================
  const handleInsertEquation = async () => {
    if (!mathfieldRef.current) return;

    const latexOriginal = mathfieldRef.current.value;
    const latexExpanded = mathfieldRef.current.getValue("latex-expanded");

    let insertContent = "";
    let isInsertingImage = false;

    let mathWidthEx = 0;
    let mathHeightEx = 0;
    let verticalAlignEx = 0; // 

    if (insertMode === "mathml") {
      // ... (Phần MathML
      
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
      // [B] CHỌN SVG CẢI TIẾN
      try {
        const mathJaxObj = (window as any).MathJax;
        if (mathJaxObj && mathJaxObj.tex2svg) {
          const svgNode = mathJaxObj.tex2svg(latexExpanded);
          const svgElement = svgNode.querySelector("svg");

          if (svgElement) {

            const FONT_SIZE_PT = 12;
            const EX_TO_PT = FONT_SIZE_PT * 0.5; // 1ex = 6pt
            const PT_TO_EMU = 12700; // 1pt = 12700 EMU

            mathWidthEx = parseFloat(svgElement.getAttribute("width") || "0");
            mathHeightEx = parseFloat(svgElement.getAttribute("height") || "0");

            const widthPt = mathWidthEx * EX_TO_PT;
            const heightPt = mathHeightEx * EX_TO_PT;
            // Render canvas ở kích thước cố định, độc lập với devicePixelRatio
            const RENDER_SCALE = 4; // chất lượng cao
            const PX_PER_PT = 96 / 72; // screen dpi / point dpi

            // MathJax thường trả về dạng: style="vertical-align: -0.25ex;"
            const styleStr = svgElement.getAttribute("style") || "";
            const vaMatch = styleStr.match(/vertical-align:\s*([-0-9.]+)ex/);
            if (vaMatch && vaMatch[1]) {
              verticalAlignEx = parseFloat(vaMatch[1]); 
            }
            // 1. Lấy kích thước chuẩn xác từ viewBox (đây là cách an toàn nhất)
            //const vb = svgElement.viewBox.baseVal;
            //const width = vb.width;
            //const height = vb.height;

            // 2. QUAN TRỌNG: Gán lại kích thước bằng pixel cho SVG
            // Nếu không có bước này, Canvas sẽ không biết kích thước ảnh là bao nhiêu
            //svgElement.setAttribute("width", width + "px");
            //svgElement.setAttribute("height", height + "px");

            const canvas = document.createElement("canvas");
            //const ctx = canvas.getContext("2d");

            // Tăng chất lượng ảnh (Scale)
            //const scale = window.devicePixelRatio || 1;
            
            //canvas.width = width * scale;
            //canvas.height = height * scale;

            canvas.width = Math.ceil(widthPt * PX_PER_PT * RENDER_SCALE);
            canvas.height = Math.ceil(heightPt * PX_PER_PT * RENDER_SCALE);
            svgElement.setAttribute("width", canvas.width + "px");
            svgElement.setAttribute("height", canvas.height + "px");
            svgElement.setAttribute("xmlns","http://www.w3.org/2000/svg");
            svgElement.setAttribute("shape-rendering","geometricPrecision");

            let svgData = new XMLSerializer().serializeToString(svgElement);
            
            svgData = svgData.replace(/currentColor/gi, "#000000");
            const utf8Encoded = encodeURIComponent(svgData).replace(/%([0-9A-F]{2})/g, (_match, p1) => {
                return String.fromCharCode(parseInt(p1, 16));
            });
            
            insertContent = window.btoa(utf8Encoded);
            isInsertingImage = true;

            //URL.revokeObjectURL(url);
          }
        }
      } catch (error) {
        console.error("Lỗi vẽ PNG:", error);
      }
    }

    // [C] GIAO TIẾP WORD (Đã fix thứ tự Sync)
    try {
      const eqId = "LSTS_" + new Date().getTime() + "_" + Math.random().toString(36).substring(2, 6);
      //Office.context.document.settings.set(eqId, latexOriginal);
      Office.context.document.settings.set(eqId, latexExpanded);
      await Office.context.document.settings.saveAsync();

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        // ---> THÊM BƯỚC ĐỌC FONT CHỮ CỦA WORD <---
        range.load("font/size");
        await context.sync();
        
        // Lấy cỡ chữ hiện tại, nếu người dùng đang bôi đen nhiều cỡ chữ khác nhau thì mặc định lấy 12
        const currentFontSize = range.font.size || 12;

        if (isInsertingImage) {
          if (!insertContent) return;

          // 1. Chèn ảnh
          const picture = range.insertInlinePictureFromBase64(insertContent, Word.InsertLocation.replace);

          // 2. PHẢI SYNC TRƯỚC KHI ĐẶT THUỘC TÍNH
          //await context.sync();

          // 3. ĐIỀU CHỈNH KÍCH THƯỚC KHỚP VỚI FONT CHỮ
          // Hệ số chuyển đổi từ "ex" sang "points" của Word
          const exToPt = currentFontSize * 0.44; 
          
          if (mathWidthEx > 0 && mathHeightEx > 0) {
              picture.width = mathWidthEx * exToPt;
              picture.height = mathHeightEx * exToPt;
          }
          picture.lockAspectRatio = true;

          // SYNC ĐỂ LƯU KÍCH THƯỚC TRƯỚC KHI XỬ LÝ VỊ TRÍ (Tùy chọn an toàn)
          //await context.sync();

          // 3. ĐIỀU CHỈNH ĐỘ CAO/THẤP CỦA ẢNH (BASELINE ALIGNMENT)
          const picRange = picture.getRange();
          if (verticalAlignEx !== 0) {
              // verticalAlignEx thường là số âm (VD: -0.25). 
              // font.position nhận giá trị pt. Số âm sẽ kéo ảnh thụt xuống dưới đường baseline.
              picRange.font.position = verticalAlignEx * exToPt - 0.1; 
          }

          // 4. Bọc Content Control
          //const picRange = picture.getRange();
          const newCc = picRange.insertContentControl();
          newCc.title = "LSTS_Math";
          newCc.tag = eqId;
          newCc.appearance = Word.ContentControlAppearance.hidden;
        } else {
          // ... (Phần MathML giữ nguyên)
          const insertedRange = range.insertHtml(insertContent, Word.InsertLocation.replace);
          const newCc = insertedRange.insertContentControl();
          newCc.title = "LSTS_Math";
          newCc.tag = eqId;
          newCc.appearance = Word.ContentControlAppearance.hidden;
        }
        await context.sync();
      });
    } catch (error) {
      console.error("Lỗi chèn vào Word:", error);
    }
  }
  // ==========================================
  // KHỞI TẠO VÀ "NGHE NGÓNG" THAO TÁC CỦA NGƯỜI DÙNG
  // ==========================================
  useEffect(() => {
    if (containerRef.current && !mathfieldRef.current) {
      const mf = new MathfieldElement();
      mf.value = "";
      mf.style.width = "100%";
      mf.style.fontSize = "22px";
      mf.style.border = "1px solid #8a8886";
      mf.style.borderRadius = "4px";
      mf.style.padding = "10px";
      mf.style.minHeight = "100px";
      mf.style.backgroundColor = "#ffffff";
      mf.style.boxSizing = "border-box";

      containerRef.current.appendChild(mf);
      mathfieldRef.current = mf;
      // ==========================================
      // NHÚNG CÁC TẬP TIN CẤU HÌNH BÊN NGOÀI VÀO MATHLIVE
      // ==========================================
      mf.inlineShortcuts = {
        ...mf.inlineShortcuts,      // Giữ lại phím tắt gốc của thư viện
        ...customInlineShortcuts,    // Đổ toàn bộ phím tắt tự chế từ file mathShortcuts.ts vào đây
        ...userDict.inline
      };

      mf.macros = {
        ...mf.macros,               // Giữ lại macro gốc
        //...customMacros,             // Đổ toàn bộ macro tự chế vào đây  
        //...userDict.macro
        ...upgradeMacros(customMacros),
        ...upgradeMacros(userDict.macro)
      };


      // Khi gõ trên hình thì code LaTeX chạy theo
      mf.addEventListener("input", () => {
        //setCurrentLatex(mf.value);
        setCurrentLatex(mf.getValue("latex-expanded"));
      });


    }

    // Lắng nghe khi bôi đen hay click chuột trong Word, gọi hàm Đọc ngược!
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


  // Hàm Lưu phím tắt mới vào bộ nhớ
  const handleSaveToDict = () => {
    if (!dictKey.trim()) {
      alert("Vui lòng nhập phím tắt!");
      return;
    }
    
    // Cập nhật State
    const updatedDict = { ...userDict };
    updatedDict[dictType][dictKey.trim()] = currentLatex; // Lấy mã LaTeX hiện tại trên màn hình
    setUserDict(updatedDict);
    
    // Lưu vào LocalStorage
    localStorage.setItem("math_user_dict", JSON.stringify(updatedDict));

    // Cập nhật nóng vào MathLive đang hiển thị
    if (mathfieldRef.current) {
      if (dictType === "inline") {
        mathfieldRef.current.inlineShortcuts = { ...mathfieldRef.current.inlineShortcuts, ...updatedDict.inline };
      } else {
        //mathfieldRef.current.macros = { ...mathfieldRef.current.macros, ...updatedDict.macro }; //Không có con trỏ
        mathfieldRef.current.macros = { ...mathfieldRef.current.macros, ...upgradeMacros(updatedDict.macro) };
      }
    }

    setIsAddDictOpen(false);
    setDictKey(""); // reset
  };

  // Hàm Xuất File JSON
  const handleExportDict = () => {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(userDict, null, 2));
    const downloadNode = document.createElement('a');
    downloadNode.setAttribute("href", dataStr);
    downloadNode.setAttribute("download", "Math_Dictionary.json");
    document.body.appendChild(downloadNode);
    downloadNode.click();
    downloadNode.remove();
  };

   // Hàm Nhập File JSON (Gộp và bảo vệ dữ liệu cũ)
  const handleImportDict = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const importedData = JSON.parse(e.target?.result as string);
        
        // Kiểm tra cấu trúc file import
        if (importedData.inline && importedData.macro) {
          
          // --- LOGIC GỘP AN TOÀN ---
          // Đặt importedData lên trước, userDict lên sau.
          // Nếu có phím tắt trùng tên, công thức cũ của người dùng sẽ được ưu tiên giữ nguyên!
          const mergedDict = {
            inline: { ...importedData.inline, ...userDict.inline },
            macro: { ...importedData.macro, ...userDict.macro }
          };

          // Lưu State và LocalStorage
          setUserDict(mergedDict);
          localStorage.setItem("math_user_dict", JSON.stringify(mergedDict));
          
          // Cập nhật nóng vào giao diện MathLive đang mở
          if (mathfieldRef.current) {
            mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...mergedDict.inline };
            //mathfieldRef.current.macros = { ...customMacros, ...mergedDict.macro };
            mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(mergedDict.macro) };
          }
          
          alert("Nhập từ điển thành công! Các phím tắt cũ của bạn vẫn được giữ nguyên an toàn.");
        } else {
          alert("Cấu trúc file JSON không đúng chuẩn từ điển (Thiếu inline hoặc macro)!");
        }
      } catch (error) {
        alert("Lỗi đọc file JSON! Hãy kiểm tra lại định dạng file.");
      }
    };
    reader.readAsText(file);
    event.target.value = ""; // Reset input để có thể import lại chính file đó nếu cần
  };

  // Mở giao diện: Chuyển Từ điển (Object) thành Danh sách (Array) để dễ bề chỉnh sửa
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

  // Hàm cập nhật (Chấp cả sửa Key lẫn Value)
  const handleItemChange = (id: string, field: "key" | "value", newValue: string) => {
    setEditItems(prev => prev.map(item => 
      item.id === id ? { ...item, [field]: newValue } : item
    ));
  };

  // Hàm xóa
  const handleItemDelete = (id: string) => {
    setEditItems(prev => prev.filter(item => item.id !== id));
  };

  // Lưu lại: Chuyển ngược từ Danh sách (Array) về lại Từ điển (Object)
  const handleSaveEditedDict = () => {
    const newDict = { inline: {} as Record<string, string>, macro: {} as Record<string, string> };
    
    editItems.forEach(item => {
      const k = item.key.trim();
      if (k) { // Bỏ qua nếu người dùng vô tình xóa trắng ô Key
        newDict[item.type][k] = item.value;
      }
    });

    setUserDict(newDict);
    localStorage.setItem("math_user_dict", JSON.stringify(newDict));
    
    if (mathfieldRef.current) {
      mathfieldRef.current.inlineShortcuts = { ...customInlineShortcuts, ...newDict.inline };
      //mathfieldRef.current.macros = { ...customMacros, ...newDict.macro };
      mathfieldRef.current.macros = { ...upgradeMacros(customMacros), ...upgradeMacros(newDict.macro) };
    }
    setIsEditDictOpen(false);
  };

  return (
    <ThemeProvider theme={officeGreenTheme}>
    <div style={{ padding: "8px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8" }}>
      <Stack tokens={{ childrenGap: 12 }}>
        {/*
        <Label style={{ fontSize: "18px", fontWeight: "bold", color: "#107c41" }}>
          {title} - Công cụ Toán học
        </Label>
        */}
        {/* HIỂN THỊ CẢNH BÁO NẾU LÀ WORD ONLINE */}
        {isOnline && (
          <MessageBar
            messageBarType={MessageBarType.blocked}
            isMultiline={true}
          >
            <strong>Phiên bản Web không được hỗ trợ chèn vào Word!</strong><br />
            Để chèn công thức, vui lòng nhấn nút <b>"Mở trong ứng dụng trên máy tính"</b> (Open in Desktop App) trên thanh công cụ của Word.
          </MessageBar>
        )}
        <div style={{ height: "175px" }}>
          <MathPalette onInsert={insertSymbol} />
        </div>

        <Label style={{ fontSize: "14px", fontWeight: "600" }}>Soạn thảo công thức:</Label>
        <div ref={containerRef} id="mathfield-container" />

        {/* Ô CODE ĐỒNG BỘ 2 CHIỀU */}
        <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
          {/*
          <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>LaTeX:</Label>
            */}
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: "4px" }}>
            <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>LaTeX:</Label>
            
            <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
              {/* Input ẩn */}
              <input type="file" accept=".json" ref={fileInputRef} style={{ display: "none" }} onChange={handleImportDict} />
              
              {/* Nút Nhập / Xuất (Có Icon, bo góc) */}
              <DefaultButton 
                iconProps={{ iconName: 'Upload' }} 
                //text="Import"
                title="Import Dictionary" 
                onClick={() => fileInputRef.current?.click()} 
                styles={{ 
                  root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, 
                  icon: { fontSize: 12 } 
                }} 
              />
              <DefaultButton 
                iconProps={{ iconName: 'Download' }} 
                //text="Export"
                title="Export Dictionary" 
                onClick={handleExportDict} 
                styles={{ 
                  root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, 
                  icon: { fontSize: 12 } 
                }} 
              />

              {/* THÊM NÚT SỬA VÀO ĐÂY */}
              <DefaultButton 
                iconProps={{ iconName: 'Edit' }} 
                //text="Edit"
                title="Edit Dictionary" 
                onClick={handleOpenEditDict} 
                styles={{ 
                  root: { height: 24, fontSize: 12, padding: "0 0px", borderRadius: "6px", minWidth: 30 }, 
                  icon: { fontSize: 12 } 
                }} 
              />

              {/* Nút Lưu Từ điển (Thêm Icon Save, bo góc, màu xanh lá đồng bộ) */}
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
            </Stack>
          </Stack>

          <textarea
            value={currentLatex}
            onChange={(e) => {
              setCurrentLatex(e.target.value);
              if (mathfieldRef.current) mathfieldRef.current.value = e.target.value;
            }}
            style={{
              width: "100%", minHeight: "50px", marginTop: "4px", padding: "6px", boxSizing: "border-box",
              fontFamily: "Consolas, monospace", fontSize: "13px", color: "#d13438",
              border: "1px solid #ccc", borderRadius: "2px", resize: "vertical"
            }}
          />
        </div>
        <Stack 
          horizontal 
          horizontalAlign="space-between" 
          verticalAlign="center" // Căn giữa nút chèn theo chiều cao của cụm bên trái
          tokens={{ childrenGap: 16 }}
          style={{ marginTop: "12px", marginBottom: "10px" }}
        >
        {/* NÚT CHỌN ĐỊNH DẠNG CHÈN */}
        
        <Stack tokens={{ childrenGap: 4 }} style={{ flexGrow: 1 }}>
          <Toggle
            label="Mode:"
            inlineLabel
            onText={isUnsupported ? "E-Editor (Only)" : "E-Editor"}
            offText="Office"
            checked={insertMode === "svg"}
            onChange={(_ev, checked) => setInsertMode(checked ? "svg" : "mathml")}
            disabled={isUnsupported} // Khóa luôn nút gạt nếu phát hiện lệnh không hỗ trợ
            styles={{ root: { marginTop: "10px", padding: "10px 0px", marginBottom: isUnsupported ? "4px" : "10px" } }}
          />
          
          {/* HIỂN THỊ CẢNH BÁO MÀU ĐỎ NẾU CHỨA LỆNH KHÔNG HỖ TRỢ */}
          {isUnsupported && (
            <Label style={{ fontSize: "12px", color: "#d13438", marginTop: "0px", marginBottom: "10px", lineHeight: "1.2" }}>
              ⚠️ Office Math is not supported. Insert E-Editor type only.
            </Label>
          )}
         
        </Stack>

        {/* NÚT CHÈN */}
        <DefaultButton
          primary
          text="Insert/Update"
          onClick={handleInsertEquation}
          style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", fontWeight: "bold", borderRadius: "8px" }}
          title="Insert new equation or update selected equation"
        />

         </Stack> 

        {/* DIALOG NHẬP TÊN TỪ ĐIỂN (Nằm ngầm, chỉ hiện khi bấm "+ Lưu vào từ điển") */}
        <Dialog
          hidden={!isAddDictOpen}
          onDismiss={() => setIsAddDictOpen(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Thêm vào từ điển',
            subText: 'Công thức này sẽ được lưu lại để bạn gõ nhanh lần sau.'
          }}
        >
          <ChoiceGroup
            options={[
              { key: 'inline', text: 'Inline (Tự động chuyển đổi khi gõ)' },
              { key: 'macro', text: 'Macro (Gõ kèm dấu \\ phía trước để gợi ý)' }
            ]}
            selectedKey={dictType}
            onChange={(_, option) => setDictType(option?.key as "inline" | "macro")}
          />
          <TextField 
            label="Shortcut key (không gõ '\' ở trước)" 
            placeholder="VD: ptb, hpt..." 
            value={dictKey} 
            onChange={(_, val) => setDictKey(val || "")} 
            styles={{ 
              root: { marginTop: 10 },
              fieldGroup: { height: 40 }, // Tăng chiều cao của khung viền lên 40px
              field: { fontSize: "15px", padding: "8px" } // Tăng nhẹ cỡ chữ để cân đối với chiều cao mới
            }}
          />
          <DialogFooter>
            <PrimaryButton onClick={handleSaveToDict} text="Lưu" />
            <DefaultButton onClick={() => setIsAddDictOpen(false)} text="Hủy" />
          </DialogFooter>
        </Dialog>
        
        {/* PANEL QUẢN LÝ VÀ TÌM KIẾM TỪ ĐIỂN */}
        <Panel
          isOpen={isEditDictOpen}
          onDismiss={() => setIsEditDictOpen(false)}
          type={PanelType.medium}
          headerText="Quản lý Từ điển"
          closeButtonAriaLabel="Đóng"
        >
          <SearchBox
            placeholder="Tìm phím tắt hoặc mã LaTeX..."
            value={searchTerm}
            onChange={(_, newValue) => setSearchTerm(newValue || "")}
            styles={{ root: { marginBottom: "16px", marginTop: "10px" } }}
          />

          <div style={{ maxHeight: "60vh", overflowY: "auto", paddingRight: "8px", overflowX: "hidden" }}>
            {["inline", "macro"].map((type) => {
              // 1. Lọc danh sách theo loại và từ khóa
              const filteredItems = editItems.filter(item => 
                item.type === type && 
                (item.key.toLowerCase().includes(searchTerm.toLowerCase()) || 
                 item.value.toLowerCase().includes(searchTerm.toLowerCase()))
              );

              if (filteredItems.length === 0) return null;

              // 2. CẮT LẤY TỐI ĐA 10 ITEM ĐỂ TỐI ƯU HIỆU SUẤT
              const displayItems = filteredItems.slice(0, 10);

              return (
                <div key={type} style={{ marginBottom: "20px" }}>
                  <Label style={{ color: "#107c41", borderBottom: "1px solid #c8c6c4", paddingBottom: "4px", marginBottom: "8px" }}>
                    {type === "inline" ? "Phím tắt gõ nhanh (Inline)" : "Lệnh Macro (\\)"}
                  </Label>
                  
                  {displayItems.map((item) => (
                    <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 8 }} key={item.id} style={{ marginBottom: "12px" }}>
                      
                      <TextField 
                        value={item.key} 
                        placeholder="Mã lệnh..."
                        onChange={(_, val) => handleItemChange(item.id, "key", val || "")}
                        styles={{ root: { width: "100px" }, field: { fontWeight: "bold" } }} 
                      />
                      
                      <TextField 
                        multiline 
                        autoAdjustHeight 
                        value={item.value} 
                        onChange={(_, val) => handleItemChange(item.id, "value", val || "")}
                        styles={{ root: { flexGrow: 1 }, field: { fontFamily: "Consolas, monospace", fontSize: "13px" } }} 
                      />
                      
                      <IconButton 
                        iconProps={{ iconName: 'Delete' }} 
                        title="Xóa phím tắt này" 
                        onClick={() => handleItemDelete(item.id)} 
                        styles={{ root: { color: "#d13438" }, rootHovered: { backgroundColor: "#fde7e9" } }} 
                      />
                    </Stack>
                  ))}

                  {/* 3. HIỂN THỊ THÔNG BÁO NẾU CÓ NHIỀU HƠN 10 KẾT QUẢ */}
                  {filteredItems.length > 10 && (
                    <Label style={{ fontStyle: "italic", color: "#d13438", fontSize: "12px", marginTop: "4px" }}>
                      * Đang hiển thị 10 / {filteredItems.length} kết quả. Hãy nhập thêm từ khóa để tìm chính xác hơn.
                    </Label>
                  )}
                </div>
              );
            })}
            
            {editItems.length === 0 && (
              <Label style={{ fontStyle: "italic", color: "#a19f9d" }}>Từ điển đang trống.</Label>
            )}
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