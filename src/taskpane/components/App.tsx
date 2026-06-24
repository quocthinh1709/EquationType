import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { DefaultButton, Stack, Label } from "@fluentui/react";
import { MathfieldElement } from "mathlive";
import "mathlive";

// Import file giao diện vừa tạo vào đây
import MathPalette from "./MathPalette";

export interface AppProps {
  title: string;
}

export default function App({ title }: AppProps) {
  const containerRef = useRef<HTMLDivElement>(null);
  const mathfieldRef = useRef<MathfieldElement | null>(null);
  const [currentLatex, setCurrentLatex] = useState("");

  useEffect(() => {
    if (containerRef.current && !mathfieldRef.current) {
      const mf = new MathfieldElement();
      mf.value = ""; 
      mf.style.width = "100%";
      mf.style.fontSize = "22px";
      mf.style.border = "1px solid #8a8886";
      mf.style.borderRadius = "4px";
      mf.style.padding = "10px";
      mf.style.minHeight = "80px";
      mf.style.backgroundColor = "#ffffff";

      mf.addEventListener("input", () => {
        setCurrentLatex(mf.value);
      });

      containerRef.current.appendChild(mf);
      mathfieldRef.current = mf;
    }
  }, []);

  // Hàm cốt lõi sẽ được truyền sang file MathPalette.tsx
  const insertSymbol = (latexCode: string) => {
    if (mathfieldRef.current) {
      mathfieldRef.current.insert(latexCode);
      mathfieldRef.current.focus();
    }
  };

  const handleInsertEquation = async () => {
    if (!mathfieldRef.current) return;
    const mathML = mathfieldRef.current.getValue("math-ml" as any);
    const htmlString = `<math xmlns="http://www.w3.org/1998/Math/MathML">${mathML}</math>`;

    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertHtml(htmlString, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (error) {
      console.error("Lỗi khi chèn công thức: ", error);
    }
  };

  return (
    <div style={{ padding: "12px", height: "100vh", boxSizing: "border-box", backgroundColor: "#faf9f8" }}>
      <Stack tokens={{ childrenGap: 12 }}>
        <Label style={{ fontSize: "18px", fontWeight: "bold", color: "#107c41" }}>
          {title} - Công cụ Toán học
        </Label>
        
        {/* KHU VỰC BẢNG KÝ HIỆU ĐƯỢC TÁCH RIÊNG */}
        {/* Truyền hàm insertSymbol vào prop onInsert để tệp con có thể gọi */}
        <MathPalette onInsert={insertSymbol} />

        {/* KHU VỰC KHUNG GÕ TOÁN */}
        <Label style={{ fontSize: "14px", fontWeight: "600" }}>Soạn thảo công thức:</Label>
        <div ref={containerRef} id="mathfield-container" />

        <div style={{ background: "#e1dfdd", padding: "8px", borderRadius: "4px" }}>
          <Label style={{ fontSize: "12px", color: "#605e5c", margin: 0 }}>Mã hệ thống (LaTeX):</Label>
          <code style={{ fontSize: "12px", wordBreak: "break-all", color: "#d13438" }}>{currentLatex || " "}</code>
        </div>

        {/* NÚT CHÈN */}
        <DefaultButton 
          primary 
          text="Chèn vào văn bản Word" 
          iconProps={{ iconName: "Add" }}
          onClick={handleInsertEquation} 
          style={{ height: "44px", backgroundColor: "#107c41", borderColor: "transparent", color: "white", fontSize: "16px", marginTop: "10px" }}
        />
      </Stack>
    </div>
  );
}