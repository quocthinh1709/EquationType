import * as React from "react";
import { DefaultButton, Stack, Pivot, PivotItem } from "@fluentui/react";

export interface MathPaletteProps {
  onInsert: (latexCode: string) => void;
}

export default function MathPalette({ onInsert }: MathPaletteProps) {
  // Tùy chỉnh CSS cho các nút
  const btnStyle = { minWidth: "42px", height: "42px", fontSize: "16px", padding: "0" };

  return (
    <Pivot aria-label="Math Palette" styles={{ root: { marginBottom: "10px" } }}>
      
      {/* ================= 1. TOÁN CƠ BẢN & TẬP HỢP ================= */}
      <PivotItem headerText="Cơ bản">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          {/* Cấu trúc cơ bản */}
          <DefaultButton text="■/□" style={btnStyle} onClick={() => onInsert("\\frac{#0}{#?}")} title="Phân số" />
          <DefaultButton text="√■" style={btnStyle} onClick={() => onInsert("\\sqrt{#0}")} title="Căn bậc 2" />
          <DefaultButton text="ⁿ√■" style={btnStyle} onClick={() => onInsert("\\sqrt[#?]{#0}")} title="Căn bậc n" />
          <DefaultButton text="■²" style={btnStyle} onClick={() => onInsert("^{#?}")} title="Lũy thừa" />
          <DefaultButton text="■_n" style={btnStyle} onClick={() => onInsert("_{#?}")} title="Chỉ số dưới" />
          <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\left|#0\\right|")} title="Trị tuyệt đối" />
          <DefaultButton text="{" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? \\\\ #? \\end{cases}")} title="Hệ phương trình" />
          
          {/* Quan hệ & Logic */}
          <DefaultButton text="±" style={btnStyle} onClick={() => onInsert("\\pm ")} title="Cộng trừ" />
          <DefaultButton text="≠" style={btnStyle} onClick={() => onInsert("\\neq ")} title="Khác" />
          <DefaultButton text="≈" style={btnStyle} onClick={() => onInsert("\\approx ")} title="Xấp xỉ" />
          <DefaultButton text="≤" style={btnStyle} onClick={() => onInsert("\\le ")} title="Nhỏ hơn hoặc bằng" />
          <DefaultButton text="≥" style={btnStyle} onClick={() => onInsert("\\ge ")} title="Lớn hơn hoặc bằng" />
          <DefaultButton text="∀" style={btnStyle} onClick={() => onInsert("\\forall ")} title="Mọi" />
          <DefaultButton text="∃" style={btnStyle} onClick={() => onInsert("\\exists ")} title="Tồn tại" />
          <DefaultButton text="∴" style={btnStyle} onClick={() => onInsert("\\therefore ")} title="Suy ra" />
          <DefaultButton text="∵" style={btnStyle} onClick={() => onInsert("\\because ")} title="Bởi vì" />

          {/* Tập hợp */}
          <DefaultButton text="∈" style={btnStyle} onClick={() => onInsert("\\in ")} title="Thuộc" />
          <DefaultButton text="∉" style={btnStyle} onClick={() => onInsert("\\notin ")} title="Không thuộc" />
          <DefaultButton text="⊂" style={btnStyle} onClick={() => onInsert("\\subset ")} title="Tập con" />
          <DefaultButton text="∪" style={btnStyle} onClick={() => onInsert("\\cup ")} title="Hợp" />
          <DefaultButton text="∩" style={btnStyle} onClick={() => onInsert("\\cap ")} title="Giao" />
          <DefaultButton text="∅" style={btnStyle} onClick={() => onInsert("\\emptyset ")} title="Rỗng" />
          <DefaultButton text="ℝ" style={btnStyle} onClick={() => onInsert("\\mathbb{R}")} title="Tập số thực" />
          <DefaultButton text="ℕ" style={btnStyle} onClick={() => onInsert("\\mathbb{N}")} title="Tập số tự nhiên" />
          <DefaultButton text="ℤ" style={btnStyle} onClick={() => onInsert("\\mathbb{Z}")} title="Tập số nguyên" />
          <DefaultButton text="ℚ" style={btnStyle} onClick={() => onInsert("\\mathbb{Q}")} title="Tập số hữu tỉ" />
        </Stack>
      </PivotItem>

      {/* ================= 2. GIẢI TÍCH & MA TRẬN ================= */}
      <PivotItem headerText="Giải tích">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          {/* Tích phân, Đạo hàm, Giới hạn */}
          <DefaultButton text="∫" style={btnStyle} onClick={() => onInsert("\\int_{#?}^{#?}")} title="Tích phân" />
          <DefaultButton text="∬" style={btnStyle} onClick={() => onInsert("\\iint")} title="Tích phân kép" />
          <DefaultButton text="∮" style={btnStyle} onClick={() => onInsert("\\oint_{#?}")} title="Tích phân đường" />
          <DefaultButton text="lim" style={btnStyle} onClick={() => onInsert("\\lim_{#? \\to #?}")} title="Giới hạn" />
          <DefaultButton text="∑" style={btnStyle} onClick={() => onInsert("\\sum_{#?}^{#?}")} title="Tổng Sigma" />
          <DefaultButton text="∏" style={btnStyle} onClick={() => onInsert("\\prod_{#?}^{#?}")} title="Tích Pi" />
          <DefaultButton text="∞" style={btnStyle} onClick={() => onInsert("\\infty")} title="Vô cực" />
          <DefaultButton text="d/dx" style={btnStyle} onClick={() => onInsert("\\frac{d}{dx} #0")} title="Đạo hàm" />
          <DefaultButton text="∂</text" style={btnStyle} onClick={() => onInsert("\\partial ")} title="Đạo hàm riêng" />
          <DefaultButton text="∇" style={btnStyle} onClick={() => onInsert("\\nabla ")} title="Toán tử gradient (Nabla)" />

          {/* Ma trận */}
          <DefaultButton text="[■]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? & #? \\\\ #? & #? \\end{bmatrix}")} title="Ma trận vuông" />
          <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\begin{vmatrix} #? & #? \\\\ #? & #? \\end{vmatrix}")} title="Định thức" />
          <DefaultButton text="⋯" style={btnStyle} onClick={() => onInsert("\\cdots ")} title="Dấu chấm giữa" />
          <DefaultButton text="⋮" style={btnStyle} onClick={() => onInsert("\\vdots ")} title="Dấu chấm dọc" />
          <DefaultButton text="⋱" style={btnStyle} onClick={() => onInsert("\\ddots ")} title="Dấu chấm chéo" />
        </Stack>
      </PivotItem>

      {/* ================= TAB: HỆ PHƯƠNG TRÌNH & MA TRẬN ================= */}
      <PivotItem headerText="Hệ phương trình & Ma trận">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          
          {/* Nhóm Hệ phương trình */}
          <DefaultButton text="{ 2 dòng" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? \\\\ #? \\end{cases}")} title="Hệ 2 phương trình" />
          <DefaultButton text="{ 3 dòng" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? \\\\ #? \\\\ #? \\end{cases}")} title="Hệ 3 phương trình" />
          <DefaultButton text="[ Hoặc" style={btnStyle} onClick={() => onInsert("\\left[ \\begin{array}{l} #? \\\\ #? \\end{array} \\right.")} title="Ngoặc vuông (Tuyển / Hoặc)" />
          <DefaultButton text="f(x)=" style={btnStyle} onClick={() => onInsert("f(x) = \\begin{cases} #?, & \\text{nếu } #? \\\\ #?, & \\text{nếu } #? \\end{cases}")} title="Hàm số cho bởi nhiều công thức" />

          {/* Nhóm Ma trận & Định thức */}
          <DefaultButton text="(■)" style={btnStyle} onClick={() => onInsert("\\begin{pmatrix} #? & #? \\\\ #? & #? \\end{pmatrix}")} title="Ma trận ngoặc tròn (2x2)" />
          <DefaultButton text="[■]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? & #? \\\\ #? & #? \\end{bmatrix}")} title="Ma trận ngoặc vuông (2x2)" />
          <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\begin{vmatrix} #? & #? \\\\ #? & #? \\end{vmatrix}")} title="Định thức (2x2)" />
          <DefaultButton text="||■||" style={btnStyle} onClick={() => onInsert("\\begin{Vmatrix} #? & #? \\\\ #? & #? \\end{Vmatrix}")} title="Chuẩn ma trận / Norm (2x2)" />
          
          {/* Nhóm Vector cột / Vector hàng */}
          <DefaultButton text="[⋮]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? \\\\ #? \\\\ #? \\end{bmatrix}")} title="Vector cột (3x1)" />
          <DefaultButton text="[⋯]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? & #? & #? \\end{bmatrix}")} title="Vector hàng (1x3)" />

          {/* Nhóm Dấu chấm hỗ trợ Ma trận tổng quát */}
          <DefaultButton text="⋯" style={btnStyle} onClick={() => onInsert("\\cdots ")} title="Chấm ngang (Cùng hàng)" />
          <DefaultButton text="⋮" style={btnStyle} onClick={() => onInsert("\\vdots ")} title="Chấm dọc (Cùng cột)" />
          <DefaultButton text="⋱" style={btnStyle} onClick={() => onInsert("\\ddots ")} title="Chấm chéo (Đường chéo chính)" />

        </Stack>
      </PivotItem>

      {/* ================= 3. HÌNH HỌC & LƯỢNG GIÁC ================= */}
      <PivotItem headerText="Hình học">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          {/* Lượng giác */}
          <DefaultButton text="sin" style={btnStyle} onClick={() => onInsert("\\sin\\left(#?\\right)")} />
          <DefaultButton text="cos" style={btnStyle} onClick={() => onInsert("\\cos\\left(#?\\right)")} />
          <DefaultButton text="tan" style={btnStyle} onClick={() => onInsert("\\tan\\left(#?\\right)")} />
          <DefaultButton text="cot" style={btnStyle} onClick={() => onInsert("\\cot\\left(#?\\right)")} />
          <DefaultButton text="arcsin" style={btnStyle} onClick={() => onInsert("\\arcsin\\left(#?\\right)")} />
          <DefaultButton text="arccos" style={btnStyle} onClick={() => onInsert("\\arccos\\left(#?\\right)")} />

          {/* Hình học */}
          <DefaultButton text="Δ" style={btnStyle} onClick={() => onInsert("\\triangle ")} title="Tam giác" />
          <DefaultButton text="∠" style={btnStyle} onClick={() => onInsert("\\angle ")} title="Góc" />
          <DefaultButton text="⊥" style={btnStyle} onClick={() => onInsert("\\perp ")} title="Vuông góc" />
          <DefaultButton text="∥" style={btnStyle} onClick={() => onInsert("\\parallel ")} title="Song song" />
          <DefaultButton text="°" style={btnStyle} onClick={() => onInsert("^{\\circ}")} title="Độ" />
          
          {/* Vector & Mũ */}
          <DefaultButton text="v⃗" style={btnStyle} onClick={() => onInsert("\\vec{#0}")} title="Vector" />
          <DefaultButton text="AB⃗" style={btnStyle} onClick={() => onInsert("\\overrightarrow{#0}")} title="Vector dài" />
          <DefaultButton text="Â" style={btnStyle} onClick={() => onInsert("\\widehat{#0}")} title="Góc trên đầu" />
          <DefaultButton text="AB⁀" style={btnStyle} onClick={() => onInsert("\\overparen{#0}")} title="Cung tròn" />
        </Stack>
      </PivotItem>

      {/* ================= 4. VẬT LÝ & ĐIỆN TỬ ================= */}
      <PivotItem headerText="Vật lý & Mạch">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          {/* Ký hiệu Vật lý */}
          <DefaultButton text="Ω" style={btnStyle} onClick={() => onInsert("\\Omega ")} title="Điện trở (Ohm)" />
          <DefaultButton text="ω" style={btnStyle} onClick={() => onInsert("\\omega ")} title="Tần số góc" />
          <DefaultButton text="λ" style={btnStyle} onClick={() => onInsert("\\lambda ")} title="Bước sóng" />
          <DefaultButton text="μ" style={btnStyle} onClick={() => onInsert("\\mu ")} title="Micro / Hệ số ma sát" />
          <DefaultButton text="Δ" style={btnStyle} onClick={() => onInsert("\\Delta ")} title="Độ biến thiên" />
          <DefaultButton text="ħ" style={btnStyle} onClick={() => onInsert("\\hbar ")} title="Hằng số Planck thu gọn" />
          <DefaultButton text="Å" style={btnStyle} onClick={() => onInsert("\\text{A}^{\\circ}")} title="Angstrom" />
          <DefaultButton text="A̅" style={btnStyle} onClick={() => onInsert("\\overline{#0}")} title="Gạch đầu (Giá trị trung bình)" />
          <DefaultButton text="A̲" style={btnStyle} onClick={() => onInsert("\\underline{#0}")} title="Gạch dưới" />

          {/* Logic Toán & Điện tử */}
          <DefaultButton text="∧" style={btnStyle} onClick={() => onInsert("\\land ")} title="Cổng AND" />
          <DefaultButton text="∨" style={btnStyle} onClick={() => onInsert("\\lor ")} title="Cổng OR" />
          <DefaultButton text="⊕" style={btnStyle} onClick={() => onInsert("\\oplus ")} title="Cổng XOR" />
          <DefaultButton text="⊗" style={btnStyle} onClick={() => onInsert("\\otimes ")} title="Phép nhân Tensor / Nhập chập" />
          <DefaultButton text="⇔" style={btnStyle} onClick={() => onInsert("\\Leftrightarrow ")} title="Tương đương" />
          <DefaultButton text="⇒" style={btnStyle} onClick={() => onInsert("\\Rightarrow ")} title="Kéo theo" />
        </Stack>
      </PivotItem>

      {/* ================= 5. HÓA HỌC & SINH HỌC ================= */}
      <PivotItem headerText="Hóa & Sinh">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          <DefaultButton text="→" style={btnStyle} onClick={() => onInsert("\\rightarrow ")} title="Phản ứng 1 chiều" />
          <DefaultButton text="⇌" style={btnStyle} onClick={() => onInsert("\\rightleftharpoons ")} title="Phản ứng thuận nghịch" />
          <DefaultButton text="↑" style={btnStyle} onClick={() => onInsert("\\uparrow ")} title="Bay hơi / Khí" />
          <DefaultButton text="↓" style={btnStyle} onClick={() => onInsert("\\downarrow ")} title="Kết tủa" />
          <DefaultButton text="t°" style={btnStyle} onClick={() => onInsert("\\xrightarrow{t^{\\circ}} ")} title="Nhiệt độ phản ứng" />
          <DefaultButton text="xt" style={btnStyle} onClick={() => onInsert("\\xrightarrow{\\text{xt}} ")} title="Xúc tác" />
          <DefaultButton text="e⁻" style={btnStyle} onClick={() => onInsert("e^{-}")} title="Electron" />
          <DefaultButton text="[A]" style={btnStyle} onClick={() => onInsert("\\left[ #0 \\right]")} title="Nồng độ Mol" />
          <DefaultButton text="A_Z X" style={btnStyle} onClick={() => onInsert("_{#?}^{#?}\\text{#?}")} title="Đồng vị nguyên tử" />
          <DefaultButton text="♀" style={btnStyle} onClick={() => onInsert("\\female ")} title="Giống cái" />
          <DefaultButton text="♂" style={btnStyle} onClick={() => onInsert("\\male ")} title="Giống đực" />
        </Stack>
      </PivotItem>

      {/* ================= 6. CHỮ CÁI HY LẠP ================= */}
      <PivotItem headerText="Hy Lạp">
        <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
          <DefaultButton text="α" style={btnStyle} onClick={() => onInsert("\\alpha ")} title="Alpha" />
          <DefaultButton text="β" style={btnStyle} onClick={() => onInsert("\\beta ")} title="Beta" />
          <DefaultButton text="γ" style={btnStyle} onClick={() => onInsert("\\gamma ")} title="Gamma" />
          <DefaultButton text="δ" style={btnStyle} onClick={() => onInsert("\\delta ")} title="Delta" />
          <DefaultButton text="ε" style={btnStyle} onClick={() => onInsert("\\varepsilon ")} title="Epsilon" />
          <DefaultButton text="θ" style={btnStyle} onClick={() => onInsert("\\theta ")} title="Theta" />
          <DefaultButton text="λ" style={btnStyle} onClick={() => onInsert("\\lambda ")} title="Lambda" />
          <DefaultButton text="π" style={btnStyle} onClick={() => onInsert("\\pi ")} title="Pi" />
          <DefaultButton text="ρ" style={btnStyle} onClick={() => onInsert("\\rho ")} title="Rho" />
          <DefaultButton text="σ" style={btnStyle} onClick={() => onInsert("\\sigma ")} title="Sigma" />
          <DefaultButton text="τ" style={btnStyle} onClick={() => onInsert("\\tau ")} title="Tau" />
          <DefaultButton text="φ" style={btnStyle} onClick={() => onInsert("\\varphi ")} title="Phi" />
          <DefaultButton text="ω" style={btnStyle} onClick={() => onInsert("\\omega ")} title="Omega" />
          {/* Chữ hoa */}
          <DefaultButton text="Γ" style={btnStyle} onClick={() => onInsert("\\Gamma ")} title="Gamma (Hoa)" />
          <DefaultButton text="Δ" style={btnStyle} onClick={() => onInsert("\\Delta ")} title="Delta (Hoa)" />
          <DefaultButton text="Θ" style={btnStyle} onClick={() => onInsert("\\Theta ")} title="Theta (Hoa)" />
          <DefaultButton text="Λ" style={btnStyle} onClick={() => onInsert("\\Lambda ")} title="Lambda (Hoa)" />
          <DefaultButton text="Π" style={btnStyle} onClick={() => onInsert("\\Pi ")} title="Pi (Hoa)" />
          <DefaultButton text="Σ" style={btnStyle} onClick={() => onInsert("\\Sigma ")} title="Sigma (Hoa)" />
          <DefaultButton text="Φ" style={btnStyle} onClick={() => onInsert("\\Phi ")} title="Phi (Hoa)" />
          <DefaultButton text="Ψ" style={btnStyle} onClick={() => onInsert("\\Psi ")} title="Psi (Hoa)" />
          <DefaultButton text="Ω" style={btnStyle} onClick={() => onInsert("\\Omega ")} title="Omega (Hoa)" />
        </Stack>
      </PivotItem>

    </Pivot>
  );
}