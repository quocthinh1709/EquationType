import * as React from "react";
import { DefaultButton, Stack, Pivot, PivotItem } from "@fluentui/react";

export interface MathPaletteProps {
    onInsert: (latexCode: string) => void;
}

export default function MathPalette({ onInsert }: MathPaletteProps) {
    // Tùy chỉnh CSS cho các nút vuông vức, gọn gàng
    const btnStyle = { minWidth: "42px", height: "42px", fontSize: "16px", padding: "0" };

    return (
        // overflowBehavior="menu" giúp các Tab thừa tự động chui vào nút 3 chấm (...)
        <Pivot
            aria-label="Math Palette"
            styles={{
                root: {
                    marginBottom: "10px",
                    overflowX: "auto",    /* Bật lại thanh cuộn ngang */
                    overflowY: "hidden",
                    whiteSpace: "nowrap",
                    paddingBottom: "8px"  /* Chừa không gian phía dưới để thanh cuộn không đè lên chữ */
                },
                link: {
                    fontSize: "14px",     /* Giảm cỡ chữ xuống 1 xíu để hiển thị được nhiều Tab hơn */
                    fontWeight: "600",
                    color: "#201f1e"
                },
                linkIsSelected: {
                    fontSize: "14px",
                    fontWeight: "bold",
                    color: "#107c41"
                }
            }}
        >

            {/* TAB 1: CƠ BẢN */}
            <PivotItem headerText="Cơ bản">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="■/□" style={btnStyle} onClick={() => onInsert("\\frac{#0}{#?}")} title="Phân số" />
                    <DefaultButton text="√■" style={btnStyle} onClick={() => onInsert("\\sqrt{#0}")} title="Căn bậc 2" />
                    <DefaultButton text="ⁿ√■" style={btnStyle} onClick={() => onInsert("\\sqrt[#?]{#0}")} title="Căn bậc n" />
                    <DefaultButton text="■²" style={btnStyle} onClick={() => onInsert("^{#?}")} title="Lũy thừa" />
                    <DefaultButton text="■_n" style={btnStyle} onClick={() => onInsert("_{#?}")} title="Chỉ số dưới" />
                    <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\left|#0\\right|")} title="Trị tuyệt đối" />
                    <DefaultButton text="±" style={btnStyle} onClick={() => onInsert("\\pm ")} title="Cộng trừ" />
                    <DefaultButton text="≠" style={btnStyle} onClick={() => onInsert("\\neq ")} title="Khác" />
                    <DefaultButton text="≈" style={btnStyle} onClick={() => onInsert("\\approx ")} title="Xấp xỉ" />
                </Stack>
            </PivotItem>

            {/* TAB 2: HỆ & MA TRẬN */}
            <PivotItem headerText="Hệ & Ma trận">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="{ 2" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? \\\\ #? \\end{cases}")} title="Hệ 2 phương trình" />
                    <DefaultButton
                        text="{ Hệ 2"
                        style={btnStyle}
                        onClick={() => onInsert("\\begin{cases} #? &#? &#? &= #? \\\\ #? &#? &#? &= #? \\end{cases}")}
                        title="Hệ 2 phương trình (Có căn chỉnh dấu =)"
                    />
                    <DefaultButton text="{ 3" style={btnStyle} onClick={() => onInsert("\\begin{cases} #? \\\\ #? \\\\ #? \\end{cases}")} title="Hệ 3 phương trình" />
                    <DefaultButton text="[ Hoặc" style={btnStyle} onClick={() => onInsert("\\left[ \\begin{array}{l} #? \\\\ #? \\end{array} \\right.")} title="Tuyển (Hoặc)" />
                    <DefaultButton text="(■)" style={btnStyle} onClick={() => onInsert("\\begin{pmatrix} #? & #? \\\\ #? & #? \\end{pmatrix}")} title="Ma trận tròn 2x2" />
                    <DefaultButton text="[■]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? & #? \\\\ #? & #? \\end{bmatrix}")} title="Ma trận vuông 2x2" />
                    <DefaultButton text="|■|" style={btnStyle} onClick={() => onInsert("\\begin{vmatrix} #? & #? \\\\ #? & #? \\end{vmatrix}")} title="Định thức 2x2" />
                    <DefaultButton text="[⋮]" style={btnStyle} onClick={() => onInsert("\\begin{bmatrix} #? \\\\ #? \\\\ #? \\end{bmatrix}")} title="Vector cột" />
                    <DefaultButton text="⋯" style={btnStyle} onClick={() => onInsert("\\cdots ")} title="Chấm ngang" />
                    <DefaultButton text="⋮" style={btnStyle} onClick={() => onInsert("\\vdots ")} title="Chấm dọc" />
                    <DefaultButton text="⋱" style={btnStyle} onClick={() => onInsert("\\ddots ")} title="Chấm chéo" />
                </Stack>
            </PivotItem>

            {/* TAB 3: GIẢI TÍCH & LƯỢNG GIÁC */}
            <PivotItem headerText="Giải tích">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="∫" style={btnStyle} onClick={() => onInsert("\\int_{#?}^{#?}")} title="Tích phân" />
                    <DefaultButton text="lim" style={btnStyle} onClick={() => onInsert("\\lim_{#? \\to #?}")} title="Giới hạn" />
                    <DefaultButton text="∑" style={btnStyle} onClick={() => onInsert("\\sum_{#?}^{#?}")} title="Tổng Sigma" />
                    <DefaultButton text="∞" style={btnStyle} onClick={() => onInsert("\\infty")} title="Vô cực" />
                    <DefaultButton text="d/dx" style={btnStyle} onClick={() => onInsert("\\frac{d}{dx} #0")} title="Đạo hàm" />
                    <DefaultButton text="sin" style={btnStyle} onClick={() => onInsert("\\sin\\left(#?\\right)")} />
                    <DefaultButton text="cos" style={btnStyle} onClick={() => onInsert("\\cos\\left(#?\\right)")} />
                    <DefaultButton text="tan" style={btnStyle} onClick={() => onInsert("\\tan\\left(#?\\right)")} />
                    <DefaultButton text="π" style={btnStyle} onClick={() => onInsert("\\pi ")} title="Pi" />
                </Stack>
            </PivotItem>

            {/* TAB 4: LÝ & ĐIỆN TỬ */}
            <PivotItem headerText="Lý & Mạch">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="v⃗" style={btnStyle} onClick={() => onInsert("\\vec{#0}")} title="Vector" />
                    <DefaultButton text="Ω" style={btnStyle} onClick={() => onInsert("\\Omega ")} title="Ohm" />
                    <DefaultButton text="ω" style={btnStyle} onClick={() => onInsert("\\omega ")} title="Tần số góc" />
                    <DefaultButton text="λ" style={btnStyle} onClick={() => onInsert("\\lambda ")} title="Bước sóng" />
                    <DefaultButton text="Δ" style={btnStyle} onClick={() => onInsert("\\Delta ")} title="Độ biến thiên" />
                    <DefaultButton text="∧" style={btnStyle} onClick={() => onInsert("\\land ")} title="Cổng AND" />
                    <DefaultButton text="∨" style={btnStyle} onClick={() => onInsert("\\lor ")} title="Cổng OR" />
                    <DefaultButton text="⊕" style={btnStyle} onClick={() => onInsert("\\oplus ")} title="Cổng XOR" />
                    <DefaultButton text="A̅" style={btnStyle} onClick={() => onInsert("\\overline{#0}")} title="Gạch đầu (NOT)" />
                </Stack>
            </PivotItem>

            {/* TAB 5: TẬP HỢP & LOGIC */}
            <PivotItem headerText="Tập hợp">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="∈" style={btnStyle} onClick={() => onInsert("\\in ")} title="Thuộc" />
                    <DefaultButton text="∉" style={btnStyle} onClick={() => onInsert("\\notin ")} title="Không thuộc" />
                    <DefaultButton text="⊂" style={btnStyle} onClick={() => onInsert("\\subset ")} title="Tập con" />
                    <DefaultButton text="∪" style={btnStyle} onClick={() => onInsert("\\cup ")} title="Hợp" />
                    <DefaultButton text="∩" style={btnStyle} onClick={() => onInsert("\\cap ")} title="Giao" />
                    <DefaultButton text="ℝ" style={btnStyle} onClick={() => onInsert("\\mathbb{R}")} title="Tập số thực" />
                    <DefaultButton text="ℕ" style={btnStyle} onClick={() => onInsert("\\mathbb{N}")} title="Tập số tự nhiên" />
                    <DefaultButton text="ℤ" style={btnStyle} onClick={() => onInsert("\\mathbb{Z}")} title="Tập số nguyên" />
                    <DefaultButton text="∀" style={btnStyle} onClick={() => onInsert("\\forall ")} title="Mọi" />
                    <DefaultButton text="∃" style={btnStyle} onClick={() => onInsert("\\exists ")} title="Tồn tại" />
                </Stack>
            </PivotItem>

            {/* TAB 6: HÓA HỌC */}
            <PivotItem headerText="Hóa học">
                <Stack horizontal wrap tokens={{ childrenGap: 6 }} style={{ marginTop: 12 }}>
                    <DefaultButton text="→" style={btnStyle} onClick={() => onInsert("\\rightarrow ")} title="Phản ứng 1 chiều" />
                    <DefaultButton text="⇌" style={btnStyle} onClick={() => onInsert("\\rightleftharpoons ")} title="Phản ứng thuận nghịch" />
                    <DefaultButton text="↑" style={btnStyle} onClick={() => onInsert("\\uparrow ")} title="Bay hơi" />
                    <DefaultButton text="↓" style={btnStyle} onClick={() => onInsert("\\downarrow ")} title="Kết tủa" />
                    <DefaultButton text="t°" style={btnStyle} onClick={() => onInsert("\\xrightarrow{t^{\\circ}} ")} title="Nhiệt độ phản ứng" />
                    <DefaultButton text="e⁻" style={btnStyle} onClick={() => onInsert("e^{-}")} title="Electron" />
                    <DefaultButton text="A_Z X" style={btnStyle} onClick={() => onInsert("_{#?}^{#?}\\text{#?}")} title="Đồng vị" />
                </Stack>
            </PivotItem>

        </Pivot>
    );
}