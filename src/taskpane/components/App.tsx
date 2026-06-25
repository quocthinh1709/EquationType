import * as React from "react";
import { useEffect, useRef, useState, useCallback } from "react";
import { DefaultButton, Stack, Label } from "@fluentui/react";
import { MathfieldElement } from "mathlive";
import "mathlive";
import MathPalette from "./MathPalette";
import { customInlineShortcuts, customMacros } from "./mathShortcuts";
// Đảm bảo load Font nội bộ chuẩn xác
MathfieldElement.fontsDirectory = "./fonts";

export interface AppProps {
  title: string;
}

export default function App({ title }: AppProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const [currentLatex, setCurrentLatex] = useState("");

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
    const mathML = mathfieldRef.current.getValue("math-ml" as any);
    const latex = mathfieldRef.current.value; 
    //const latex = mathfieldRef.current.getValue("latex-expanded");
    const htmlString = `<math xmlns="http://www.w3.org/1998/Math/MathML">${mathML}</math>`;


    try {
      // Tạo ID duy nhất và lưu mã LaTeX dài vô tận vào Settings
      const eqId = "LSTS_" + new Date().getTime();
      Office.context.document.settings.set(eqId, latex);
      Office.context.document.settings.saveAsync();

      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const currentCc = range.parentContentControlOrNullObject;
        currentCc.load("title");
        await context.sync();

        let targetRange = range;

        // TÍNH NĂNG GHI ĐÈ: Nếu đang bôi đen/đứng ở công thức cũ, chọn nó để xóa
        if (!currentCc.isNullObject && currentCc.title === "LSTS_Math") {
          targetRange = currentCc.getRange("Whole");
        }

        // Chèn hình ảnh toán học mới
        const insertedRange = targetRange.insertHtml(htmlString, Word.InsertLocation.replace);
        
        // Bọc lại bằng khung tàng hình để lần sau còn bấm vào sửa được
        const newCc = insertedRange.insertContentControl();
        newCc.title = "LSTS_Math";
        newCc.tag = eqId; 
        newCc.appearance = Word.ContentControlAppearance.hidden;

        await context.sync();
      });
    } catch (error) {
      console.error("Lỗi khi chèn công thức: ", error);
    }
  };

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

    // Lắng nghe: Cứ mỗi khi anh bôi đen hay click chuột trong Word, gọi hàm Đọc ngược!
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
    <div title={title} style={{ padding: "12px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8" }}>
      <Stack tokens={{ childrenGap: 12 }}>
        
        <Label style={{ fontSize: "18px", fontWeight: "bold", color: "#107c41" }}>
          {title} - Công cụ Toán học
        </Label>
        
        <MathPalette onInsert={insertSymbol} />

        <Label style={{ fontSize: "14px", fontWeight: "600" }}>Soạn thảo công thức:</Label>
        <div ref={containerRef} id="mathfield-container" />

        {/* Ô CODE ĐỒNG BỘ 2 CHIỀU */}
        <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
          <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>Mã hệ thống (LaTeX):</Label>
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

        {/* NÚT CHÈN MẠNH MẼ */}
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