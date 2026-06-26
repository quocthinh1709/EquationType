import * as React from "react";
import { DefaultButton, Stack, Pivot, PivotItem } from "@fluentui/react";

declare global {
  namespace JSX {
    interface IntrinsicElements {
      "math-field": React.DetailedHTMLProps<React.HTMLAttributes<HTMLElement>, HTMLElement> & {
        "read-only"?: boolean | string;
      };
    }
  }
}
// Hàm tạo nút bấm chứa công thức hiển thị trực quan
// Hàm tạo nút bấm chứa công thức hiển thị trực quan (Đã fix lỗi click và kích thước)
const MathBtn = ({ displayLatex, insertLatex, title, onInsert }: { displayLatex: string, insertLatex: string, title: string, onInsert: (l: string) => void }) => {
    
    // 1. Đồng bộ kích thước chuẩn 42px như tất cả các nút DefaultButton khác
    const btnStyle = { 
        minWidth: "42px", 
        height: "42px", 
        padding: "0", 
        position: "relative" as any // Cần relative để làm chuẩn cho lớp kính tàng hình
    }; 
    
    return (
        <DefaultButton style={btnStyle} onClick={() => onInsert(insertLatex)} title={title}>
            
            {/* 2. LỚP KÍNH TÀNG HÌNH: Hứng trọn cú click chuột, bảo vệ không cho math-field nuốt sự kiện */}
            <div 
                style={{ 
                    position: "absolute", 
                    top: 0, left: 0, right: 0, bottom: 0, 
                    zIndex: 10, 
                    cursor: "pointer" 
                }} 
            />
            
            {/* 3. Khung hiển thị công thức */}
            <math-field 
                read-only 
                style={{ 
                    pointerEvents: "none", 
                    border: "none", 
                    background: "transparent", 
                    padding: 0, 
                    margin: 0,
                    minHeight: "auto",
                    fontSize: "18px", // Ép cỡ chữ nhỏ lại một chút để vừa khít ô 42px
                    outline: "none"
                }}
            >
                {displayLatex}
            </math-field>

        </DefaultButton>
    );
};




export interface MathPaletteProps {
    onInsert: (latexCode: string) => void;
}

export default function MathPalette({ onInsert }: MathPaletteProps) {
    // Kích thước chuẩn cho nút bấm để vừa vặn các biểu tượng lớn
    const btnStyle = { minWidth: "42px", height: "42px", fontSize: "16px", padding: "0" };

    return (
        <Pivot
            aria-label="Math Palette Full"
            styles={{
                root: {
                    marginBottom: "10px",
                    overflowX: "auto",
                    overflowY: "hidden",
                    whiteSpace: "nowrap",
                    paddingBottom: "8px"
                },
                link: { fontSize: "14px", fontWeight: "600", color: "#201f1e" },
                linkIsSelected: { fontSize: "14px", fontWeight: "bold", color: "#107c41" }
            }}
        >

            {/* TAB 1: ĐẠI SỐ CƠ BẢN */}
            <PivotItem headerText="Đại số">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="■/□" style={btnStyle} onClick={() => onInsert("\\frac{#0}{#?}")} title="Phân số" />
                    <DefaultButton text="√■" style={btnStyle} onClick={() => onInsert("\\sqrt{#0}")} title="Căn bậc 2" />
                    <DefaultButton text="ⁿ√■" style={btnStyle} onClick={() => onInsert("\\sqrt[#?]{#0}")} title="Căn bậc n" />
                    <DefaultButton text="■²" style={btnStyle} onClick={() => onInsert("^{#?}")} title="Lũy thừa / Superscript" />
                    <DefaultButton text="■_n" style={btnStyle} onClick={() => onInsert("_{#?}")} title="Chỉ số dưới / Subscript" />
                    <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\left|#0\\right|")} title="Trị tuyệt đối" />
                    <DefaultButton text="±" style={btnStyle} onClick={() => onInsert("\\pm ")} title="Cộng trừ" />
                    <DefaultButton text="∓" style={btnStyle} onClick={() => onInsert("\\mp ")} title="Trừ cộng" />
                    <DefaultButton text="×" style={btnStyle} onClick={() => onInsert("\\times ")} title="Dấu nhân" />
                    <DefaultButton text="÷" style={btnStyle} onClick={() => onInsert("\\div ")} title="Dấu chia" />
                    <DefaultButton text="·" style={btnStyle} onClick={() => onInsert("\\cdot ")} title="Dấu chấm nhân" />
                    <DefaultButton text="A̅" style={btnStyle} onClick={() => onInsert("\\overline{#0}")} title="Gạch đầu (Overline)" />
                    <DefaultButton text="A̲" style={btnStyle} onClick={() => onInsert("\\underline{#0}")} title="Gạch dưới (Underline)" />
                </Stack>
            </PivotItem>

            {/* TAB 2: TOÁN TỬ & LOGIC */}
            <PivotItem headerText="Toán tử & Logic">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="≠" style={btnStyle} onClick={() => onInsert("\\neq ")} title="Khác" />
                    <DefaultButton text="≈" style={btnStyle} onClick={() => onInsert("\\approx ")} title="Xấp xỉ" />
                    <DefaultButton text="≡" style={btnStyle} onClick={() => onInsert("\\equiv ")} title="Đồng dư / Tương đương" />
                    <DefaultButton text="≤" style={btnStyle} onClick={() => onInsert("\\le ")} title="Nhỏ hơn hoặc bằng" />
                    <DefaultButton text="≥" style={btnStyle} onClick={() => onInsert("\\ge ")} title="Lớn hơn hoặc bằng" />
                    <DefaultButton text="∀" style={btnStyle} onClick={() => onInsert("\\forall ")} title="Mọi" />
                    <DefaultButton text="∃" style={btnStyle} onClick={() => onInsert("\\exists ")} title="Tồn tại" />
                    <DefaultButton text="∄" style={btnStyle} onClick={() => onInsert("\\nexists ")} title="Không tồn tại" />
                    <DefaultButton text="∴" style={btnStyle} onClick={() => onInsert("\\therefore ")} title="Suy ra (Kết luận)" />
                    <DefaultButton text="∵" style={btnStyle} onClick={() => onInsert("\\because ")} title="Bởi vì (Giải thích)" />
                    <DefaultButton text="∧" style={btnStyle} onClick={() => onInsert("\\land ")} title="Và (AND)" />
                    <DefaultButton text="∨" style={btnStyle} onClick={() => onInsert("\\lor ")} title="Hoặc (OR)" />
                    <DefaultButton text="¬" style={btnStyle} onClick={() => onInsert("\\neg ")} title="Phủ định (NOT)" />
                </Stack>
            </PivotItem>

            {/* TAB 3: TẬP HỢP */}
            <PivotItem headerText="Tập hợp">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="∈" style={btnStyle} onClick={() => onInsert("\\in ")} title="Thuộc" />
                    <DefaultButton text="∉" style={btnStyle} onClick={() => onInsert("\\notin ")} title="Không thuộc" />
                    <DefaultButton text="⊂" style={btnStyle} onClick={() => onInsert("\\subset ")} title="Tập con thực sự" />
                    <DefaultButton text="⊆" style={btnStyle} onClick={() => onInsert("\\subseteq ")} title="Tập con hoặc bằng" />
                    <DefaultButton text="⊃" style={btnStyle} onClick={() => onInsert("\\supset ")} title="Chứa tập con" />
                    <DefaultButton text="∪" style={btnStyle} onClick={() => onInsert("\\cup ")} title="Hợp" />
                    <DefaultButton text="∩" style={btnStyle} onClick={() => onInsert("\\cap ")} title="Giao" />
                    <DefaultButton text="∖" style={btnStyle} onClick={() => onInsert("\\setminus ")} title="Hiệu hai tập hợp" />
                    <DefaultButton text="∅" style={btnStyle} onClick={() => onInsert("\\emptyset ")} title="Tập rỗng" />
                    <DefaultButton text="ℝ" style={btnStyle} onClick={() => onInsert("\\mathbb{R}")} title="Tập số thực" />
                    <DefaultButton text="ℕ" style={btnStyle} onClick={() => onInsert("\\mathbb{N}")} title="Tập số tự nhiên" />
                    <DefaultButton text="ℤ" style={btnStyle} onClick={() => onInsert("\\mathbb{Z}")} title="Tập số nguyên" />
                    <DefaultButton text="ℚ" style={btnStyle} onClick={() => onInsert("\\mathbb{Q}")} title="Tập số hữu tỉ" />
                </Stack>
            </PivotItem>

            {/* TAB 4: GIẢI TÍCH */}
            <PivotItem headerText="Giải tích">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="∫" style={btnStyle} onClick={() => onInsert("\\int_{#0}^{#1}")} title="Tích phân" />
                    <DefaultButton text="∬" style={btnStyle} onClick={() => onInsert("\\iint ")} title="Tích phân kép" />
                    <DefaultButton text="∮" style={btnStyle} onClick={() => onInsert("\\oint ")} title="Tích phân đường" />
                    <DefaultButton text="lim" style={btnStyle} onClick={() => onInsert("\\lim_{#? \\to #?}")} title="Giới hạn" />
                    <DefaultButton text="∑" style={btnStyle} onClick={() => onInsert("\\sum_{#?}^{#?}")} title="Tổng Sigma" />
                    <DefaultButton text="∏" style={btnStyle} onClick={() => onInsert("\\prod_{#?}^{#?}")} title="Tích Pi" />
                    <DefaultButton text="∞" style={btnStyle} onClick={() => onInsert("\\infty")} title="Vô cực" />
                    <DefaultButton text="d/dx" style={btnStyle} onClick={() => onInsert("\\frac{d}{dx} #0")} title="Đạo hàm" />
                    <DefaultButton text="∂</text" style={btnStyle} onClick={() => onInsert("\\partial ")} title="Đạo hàm riêng" />
                    <DefaultButton text="∇" style={btnStyle} onClick={() => onInsert("\\nabla ")} title="Nabla / Gradient" />
                    <DefaultButton text="!" style={btnStyle} onClick={() => onInsert("! ")} title="Giai thừa" />
                </Stack>
            </PivotItem>

            {/* TAB 5: HỆ & MA TRẬN */}
            <PivotItem headerText="Hệ & Ma trận">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="{ Hệ 2" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? &= #? \\\\ #? &= #? \\end{cases}")} title="Hệ 2 phương trình (căn chỉnh thẳng =)" />
                    <DefaultButton text="{ Hệ 3" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? &= #? \\\\ #? &= #? \\\\ #? &= #? \\end{cases}")} title="Hệ 3 phương trình (căn chỉnh thẳng =)" />
                    <DefaultButton text="[ Hoặc" style={btnStyle} onClick={() => onInsert("\\left[ \\begin{array}{l} #? \\\\ #? \\end{array} \\right.")} title="Ngoặc vuông (Hoặc / Tuyển)" />
                    <DefaultButton text="f(x)" style={btnStyle} onClick={() => onInsert("f(x) = \\begin{cases} #?, & \\text{nếu } #? \\\\ #?, & \\text{nếu } #? \\end{cases}")} title="Hàm số cho bởi nhiều công thức" />
                    <DefaultButton text="(■)" style={btnStyle} onClick={() => onInsert("\\begin{pmatrix} #? & #? \\\\ #? & #? \\end{pmatrix}")} title="Ma trận tròn 2x2" />
                    <DefaultButton text="[■]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? & #? \\\\ #? & #? \\end{bmatrix}")} title="Ma trận vuông 2x2" />
                    <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\begin{vmatrix} #? & #? \\\\ #? & #? \\end{vmatrix}")} title="Định thức 2x2" />
                    <DefaultButton text="[⋮]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? \\\\ #? \\\\ #? \\end{bmatrix}")} title="Vector cột 3x1" />
                    <DefaultButton text="⋯" style={btnStyle} onClick={() => onInsert("\\cdots ")} title="Chấm ngang giữa" />
                    <DefaultButton text="⋮" style={btnStyle} onClick={() => onInsert("\\vdots ")} title="Chấm dọc" />
                    <DefaultButton text="⋱" style={btnStyle} onClick={() => onInsert("\\ddots ")} title="Chấm chéo" />
                </Stack>
            </PivotItem>

            {/* TAB 6: HÌNH & LƯỢNG GIÁC */}
            <PivotItem headerText="Hình & Lượng giác">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="sin" style={btnStyle} onClick={() => onInsert("\\sin\\left(#?\\right)")} />
                    <DefaultButton text="cos" style={btnStyle} onClick={() => onInsert("\\cos\\left(#?\\right)")} />
                    <DefaultButton text="tan" style={btnStyle} onClick={() => onInsert("\\tan\\left(#?\\right)")} />
                    <DefaultButton text="cot" style={btnStyle} onClick={() => onInsert("\\cot\\left(#?\\right)")} />
                    <DefaultButton text="π" style={btnStyle} onClick={() => onInsert("\\pi ")} title="Pi" />
                    <DefaultButton text="°" style={btnStyle} onClick={() => onInsert("^{\\circ}")} title="Độ" />
                    <DefaultButton text="Δ" style={btnStyle} onClick={() => onInsert("\\triangle ")} title="Tam giác" />
                    <DefaultButton text="∠" style={btnStyle} onClick={() => onInsert("\\angle ")} title="Góc" />
                    <DefaultButton text="⊥" style={btnStyle} onClick={() => onInsert("\\perp ")} title="Vuông góc" />
                    <DefaultButton text="∥" style={btnStyle} onClick={() => onInsert("\\parallel ")} title="Song song" />
                    
                    <MathBtn displayLatex="\vec{#?}" insertLatex="\vec{#0}" title="Vector" onInsert={onInsert} />
                    <DefaultButton text="Â" style={btnStyle} onClick={() => onInsert("\\hat{#0}")} title="Góc trên đầu (ABC)" />
                    <MathBtn displayLatex="\overset{\frown}{#0}" insertLatex="\overset{\frown}{#0}" title="Cung tròn" onInsert={onInsert} />
                </Stack>
            </PivotItem>

            {/* TAB 7: LÝ & HÓA HỌC */}
            <PivotItem headerText="Lý & Hóa học">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="→" style={btnStyle} onClick={() => onInsert("\\rightarrow ")} title="Phản ứng 1 chiều" />
                    <DefaultButton text="⇌" style={btnStyle} onClick={() => onInsert("\\rightleftharpoons ")} title="Phản ứng thuận nghịch" />
                    <DefaultButton text="↑" style={btnStyle} onClick={() => onInsert("\\uparrow ")} title="Bay hơi" />
                    <DefaultButton text="↓" style={btnStyle} onClick={() => onInsert("\\downarrow ")} title="Kết tủa" />
                    <DefaultButton text="t°" style={btnStyle} onClick={() => onInsert("\\xrightarrow{t^{\\circ}} ")} title="Phản ứng cần nhiệt độ" />
                    <DefaultButton text="xt" style={btnStyle} onClick={() => onInsert("\\xrightarrow{\\text{xt}} ")} title="Phản ứng cần xúc tác" />
                    <DefaultButton text="e⁻" style={btnStyle} onClick={() => onInsert("e^{-}")} title="Electron" />
                    <DefaultButton text="A_Z X" style={btnStyle} onClick={() => onInsert("\\ce{^{#?}_{#?}#?}")} title="Ký hiệu nguyên tử / Đồng vị" />
                    <DefaultButton text="Ω" style={btnStyle} onClick={() => onInsert("\\Omega ")} title="Ohm" />
                    <DefaultButton text="ω" style={btnStyle} onClick={() => onInsert("\\omega ")} title="Tần số góc (Omega)" />
                    <DefaultButton text="λ" style={btnStyle} onClick={() => onInsert("\\lambda ")} title="Bước sóng (Lambda)" />
                    <DefaultButton text="μ" style={btnStyle} onClick={() => onInsert("\\mu ")} title="Micro / Hệ số ma sát" />
                    <DefaultButton text="Å" style={btnStyle} onClick={() => onInsert("\\text{\\AA}")} title="Angstrom" />
                    <DefaultButton text="ħ" style={btnStyle} onClick={() => onInsert("\\hbar ")} title="Hằng số Planck thu gọn" />
                </Stack>
            </PivotItem>

            {/* TAB 8: MŨI TÊN */}
            <PivotItem headerText="Mũi tên">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="←" style={btnStyle} onClick={() => onInsert("\\leftarrow ")} />
                    <DefaultButton text="→" style={btnStyle} onClick={() => onInsert("\\rightarrow ")} />
                    <DefaultButton text="↔" style={btnStyle} onClick={() => onInsert("\\leftrightarrow ")} />
                    <DefaultButton text="⇐" style={btnStyle} onClick={() => onInsert("\\Leftarrow ")} />
                    <DefaultButton text="⇒" style={btnStyle} onClick={() => onInsert("\\Rightarrow ")} title="Kéo theo / Implies" />
                    <DefaultButton text="⇔" style={btnStyle} onClick={() => onInsert("\\Leftrightarrow ")} title="Tương đương" />
                    <DefaultButton text="↦" style={btnStyle} onClick={() => onInsert("\\mapsto ")} title="Ánh xạ" />
                    <DefaultButton text="↗" style={btnStyle} onClick={() => onInsert("\\nearrow ")} />
                    <DefaultButton text="↘" style={btnStyle} onClick={() => onInsert("\\searrow ")} />
                    <DefaultButton text="↙" style={btnStyle} onClick={() => onInsert("\\swarrow ")} />
                    <DefaultButton text="↖" style={btnStyle} onClick={() => onInsert("\\nwarrow ")} />
                </Stack>
            </PivotItem>

            {/* TAB 9: HY LẠP */}
            <PivotItem headerText="Hy Lạp">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    {/* Chữ thường */}
                    <DefaultButton text="α" style={btnStyle} onClick={() => onInsert("\\alpha ")} title="Alpha" />
                    <DefaultButton text="β" style={btnStyle} onClick={() => onInsert("\\beta ")} title="Beta" />
                    <DefaultButton text="γ" style={btnStyle} onClick={() => onInsert("\\gamma ")} title="Gamma" />
                    <DefaultButton text="δ" style={btnStyle} onClick={() => onInsert("\\delta ")} title="Delta" />
                    <DefaultButton text="ε" style={btnStyle} onClick={() => onInsert("\\varepsilon ")} title="Epsilon" />
                    <DefaultButton text="θ" style={btnStyle} onClick={() => onInsert("\\theta ")} title="Theta" />
                    <DefaultButton text="λ" style={btnStyle} onClick={() => onInsert("\\lambda ")} title="Lambda" />
                    <DefaultButton text="μ" style={btnStyle} onClick={() => onInsert("\\mu ")} title="Mu" />
                    <DefaultButton text="ν" style={btnStyle} onClick={() => onInsert("\\nu ")} title="Nu" />
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