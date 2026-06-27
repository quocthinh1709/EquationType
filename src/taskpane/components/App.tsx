import * as React from "react";
import { useEffect, useRef, useState, useCallback } from "react";
//import { DefaultButton, Stack, Label } from "@fluentui/react";
import { DefaultButton, Stack, Label, ChoiceGroup, IChoiceGroupOption, Toggle } from "@fluentui/react";

import { MathfieldElement } from "mathlive";
import "mathlive";
import MathPalette from "./MathPalette";
import { customInlineShortcuts, customMacros } from "./mathShortcuts";
import { unsupportedMathMlCommands } from "./unsupportedCommands";

// Đảm bảo load Font nội bộ chuẩn xác
MathfieldElement.fontsDirectory = "./fonts";

export interface AppProps {
  title: string;
}

export default function App({ title }: AppProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const [currentLatex, setCurrentLatex] = useState("");

  // THÊM BIẾN TRẠNG THÁI LƯU CHẾ ĐỘ CHÈN
  const [insertMode, setInsertMode] = useState<string>("mathml");

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

            mathWidthEx = parseFloat(svgElement.getAttribute("width") || "0");
            mathHeightEx = parseFloat(svgElement.getAttribute("height") || "0");

            // MathJax thường trả về dạng: style="vertical-align: -0.25ex;"
            const styleStr = svgElement.getAttribute("style") || "";
            const vaMatch = styleStr.match(/vertical-align:\s*([-0-9.]+)ex/);
            if (vaMatch && vaMatch[1]) {
              verticalAlignEx = parseFloat(vaMatch[1]); 
            }
            // 1. Lấy kích thước chuẩn xác từ viewBox (đây là cách an toàn nhất)
            const vb = svgElement.viewBox.baseVal;
            const width = vb.width;
            const height = vb.height;

            // 2. QUAN TRỌNG: Gán lại kích thước bằng pixel cho SVG
            // Nếu không có bước này, Canvas sẽ không biết kích thước ảnh là bao nhiêu
            svgElement.setAttribute("width", width + "px");
            svgElement.setAttribute("height", height + "px");

            const canvas = document.createElement("canvas");
            const ctx = canvas.getContext("2d");

            // Tăng chất lượng ảnh (Scale)
            const scale = window.devicePixelRatio || 1;
            canvas.width = width * scale;
            canvas.height = height * scale;

            const svgData = new XMLSerializer().serializeToString(svgElement);
            const img = new Image();
            const svgBlob = new Blob([svgData], { type: "image/svg+xml;charset=utf-8" });
            const url = URL.createObjectURL(svgBlob);

            // ... (Phần còn lại giữ nguyên như cũ)
            await new Promise((resolve, reject) => {
              img.onload = resolve;
              img.onerror = reject;
              img.src = url;
            });

            ctx?.scale(scale, scale);

            ctx?.drawImage(img, 0, 0);

            insertContent = canvas.toDataURL("image/png").split(",")[1];
            isInsertingImage = true;
            URL.revokeObjectURL(url);
          }
        }
      } catch (error) {
        console.error("Lỗi vẽ PNG:", error);
      }
    }

    // [C] GIAO TIẾP WORD (Đã fix thứ tự Sync)
    try {
      const eqId = "LSTS_" + new Date().getTime();
      Office.context.document.settings.set(eqId, latexOriginal);
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
          const exToPt = currentFontSize * 0.5; 
          
          if (mathWidthEx > 0 && mathHeightEx > 0) {
              picture.width = mathWidthEx * exToPt;
              picture.height = mathHeightEx * exToPt;
          }
          picture.lockAspectRatio = true;

          // 3. ĐIỀU CHỈNH ĐỘ CAO/THẤP CỦA ẢNH (BASELINE ALIGNMENT)
          const picRange = picture.getRange();
          if (verticalAlignEx !== 0) {
              // verticalAlignEx thường là số âm (VD: -0.25). 
              // font.position nhận giá trị pt. Số âm sẽ kéo ảnh thụt xuống dưới đường baseline.
              picRange.font.position = verticalAlignEx * exToPt; 
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
        ...customInlineShortcuts    // Đổ toàn bộ phím tắt tự chế từ file mathShortcuts.ts vào đây
      };

      mf.macros = {
        ...mf.macros,               // Giữ lại macro gốc
        ...customMacros             // Đổ toàn bộ macro tự chế vào đây
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

  return (
    <div title={title} style={{ padding: "8px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8" }}>
      <Stack tokens={{ childrenGap: 12 }}>
        {/*
        <Label style={{ fontSize: "18px", fontWeight: "bold", color: "#107c41" }}>
          {title} - Công cụ Toán học
        </Label>
        */}
        <MathPalette onInsert={insertSymbol} />

        <Label style={{ fontSize: "14px", fontWeight: "600" }}>Soạn thảo công thức:</Label>
        <div ref={containerRef} id="mathfield-container" />

        {/* Ô CODE ĐỒNG BỘ 2 CHIỀU */}
        <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
          <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>LaTeX:</Label>
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

        {/* NÚT CHỌN ĐỊNH DẠNG CHÈN */}
        <div>
          <Toggle
            label="Mode:"
            inlineLabel
            onText={isUnsupported ? "Image (Only)" : "Image"}
            offText="MathML"
            checked={insertMode === "svg"}
            onChange={(_ev, checked) => setInsertMode(checked ? "svg" : "mathml")}
            disabled={isUnsupported} // Khóa luôn nút gạt nếu phát hiện lệnh không hỗ trợ
            styles={{ root: { marginTop: "10px", marginBottom: isUnsupported ? "4px" : "10px" } }}
          />
          
          {/* HIỂN THỊ CẢNH BÁO MÀU ĐỎ NẾU CHỨA LỆNH KHÔNG HỖ TRỢ */}
          {isUnsupported && (
            <Label style={{ fontSize: "12px", color: "#d13438", marginTop: "0px", marginBottom: "10px", lineHeight: "1.2" }}>
              ⚠️ Công thức này không được Office Math hỗ trợ. Chế độ chèn ảnh sẽ được sử dụng.
            </Label>
          )}
        </div>

        {/* NÚT CHÈN */}
        <DefaultButton
          primary
          text="Chèn / Cập nhật"
          onClick={handleInsertEquation}
          style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", fontWeight: "bold" }}
          title="Chèn công thức mới hoặc ghi đè công thức đang bôi đen"
        />

      </Stack>
    </div>
  );
}