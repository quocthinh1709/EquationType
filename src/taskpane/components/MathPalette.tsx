import * as React from "react";
import { DefaultButton, Stack, Pivot, PivotItem } from "@fluentui/react";

export const PaletteColors = {
    office: "#107c41", // 
    latex: "#f5b833", // Vàng chú ý
    default: undefined  // Để mặc định của hệ thống
};

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
const MathBtn = ({
    displayLatex,
    insertLatex,
    title,
    onInsert,
    borderColor // Không để giá trị mặc định ở đây để ưu tiên dùng style mặc định
}: {
    displayLatex: string,
    insertLatex: string,
    title: string,
    onInsert: (l: string) => void,
    borderColor?: string
}) => {

    // Nếu không truyền borderColor, border sẽ là null (hoặc bỏ qua) để dùng kiểu dáng mặc định của DefaultButton
    const btnStyle = {
        minWidth: "40px",
        height: "40px",
        padding: "0",
        position: "relative" as any,
        // Nếu có borderColor thì mới vẽ viền, không thì để mặc định
        border: borderColor ? `1.4px solid ${borderColor}` : undefined,
        backgroundColor: "transparent",
        borderRadius: "5px"
    };

    return (
        <DefaultButton
            style={btnStyle}
            onClick={() => onInsert(insertLatex)}
            title={title}
            onRenderDescription={() => null}
        >
            <div
                style={{
                    position: "absolute",
                    top: 0, left: 0, right: 0, bottom: 0,
                    zIndex: 10,
                    cursor: "pointer"
                }}
            />

            <math-field
                read-only
                style={{
                    pointerEvents: "none",
                    border: "none",
                    background: "transparent",
                    padding: 0,
                    margin: 0,
                    minHeight: "auto",
                    fontSize: "14px",
                    fontWeight: "bold",
                    outline: "none",
                    // Màu chữ cố định là màu mặc định (không ép theo màu viền nữa)
                    color: "inherit"
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
    const pivotWrapperRef = React.useRef<HTMLDivElement>(null);
    const scrollTabs = (direction: 'left' | 'right') => {
        if (pivotWrapperRef.current) {
            const tabList = pivotWrapperRef.current.querySelector('[role="tablist"]');
            if (tabList) {
                const scrollAmount = 250;
                (tabList as HTMLElement).scrollBy({
                    left: direction === 'left' ? -scrollAmount : scrollAmount,
                    behavior: 'smooth'
                });
            }
        }
    };

    return (
        // =====================================
        // 2. KHUNG BAO NGOÀI (Đã gỡ bỏ onWheel)
        // =====================================
        <div ref={pivotWrapperRef} style={{ position: "relative", marginBottom: "10px" }}>

            {/* NÚT BÊN TRÁI */}
            <DefaultButton
                onClick={() => scrollTabs('left')}
                style={{
                    position: "absolute", top: 0, left: 0, zIndex: 10,
                    minWidth: "32px", width: "32px", height: "44px", padding: 0,
                    border: "none", backgroundColor: "#faf9f8", color: "#107c41", fontWeight: "bold"
                }}
                title="Cuộn sang trái"
            >
                ❮
            </DefaultButton>

            {/* NÚT BÊN PHẢI */}
            <DefaultButton
                onClick={() => scrollTabs('right')}
                style={{
                    position: "absolute", top: 0, right: 0, zIndex: 10,
                    minWidth: "32px", width: "32px", height: "44px", padding: 0,
                    border: "none", backgroundColor: "#faf9f8", color: "#107c41", fontWeight: "bold"
                }}
                title="Cuộn sang phải"
            >
                ❯
            </DefaultButton>

            {/* ===================================== */}
            {/* 3. KHUNG PIVOT (Dùng cuộn nguyên bản của trình duyệt) */}
            {/* ===================================== */}
            <Pivot
                aria-label="Math Palette Full"
                styles={{
                    root: {
                        overflowX: "auto", // Mở khóa tính năng cuộn nguyên bản (Vuốt Touchpad, Shift + Cuộn)
                        whiteSpace: "nowrap",
                        paddingLeft: "32px", // Chừa chỗ cho nút trái
                        paddingRight: "32px", // Chừa chỗ cho nút phải

                        // Tàng hình thanh trượt xám xịt
                        scrollbarWidth: "none",
                        msOverflowStyle: "none",
                        selectors: {
                            "&::-webkit-scrollbar": { display: "none" }
                        }
                    },
                    itemContainer: {
                        whiteSpace: "normal", // Đảm bảo các nút ở dưới không bị lỗi ngang
                        paddingTop: "5px"
                    },
                    link: { fontSize: "14px", fontWeight: "600", color: "#201f1e" },
                    linkIsSelected: { fontSize: "14px", fontWeight: "bold", color: "#107c41" }
                }}
            >

                {/* TAB 1: ĐẠI SỐ CƠ BẢN */}
                <PivotItem headerText="Đại số">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\frac{#0}{#?}" insertLatex="\frac{#0}{#?}" title="\frac" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}/{#?}" insertLatex="{#0}/{#?}" title="/" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt{#0}" insertLatex="\sqrt{#0}" title="\sqrt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt[#?]{#0}" insertLatex="\sqrt[#?]{#0}" title="\sqrt[n]" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#?}^{#?}" insertLatex="{#0}^{#?}" title="^{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}_{#?}" insertLatex="{#0}_{#?}" title="_{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\left|#0\right|" insertLatex="\left|#0\right|" title="\left|\right|" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\pm" insertLatex="\pm " title="\pm" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mp" insertLatex="\mp " title="\mp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\times" insertLatex="\times " title="\times" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\div" insertLatex="\div " title="\div" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdot" insertLatex="\cdot " title="\cdot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overline{#0}" insertLatex="\overline{#0}" title="\overline" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\underline{#0}" insertLatex="\underline{#0}" title="\underline" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 2: TOÁN TỬ & LOGIC */}
                <PivotItem headerText="Toán tử & Logic">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\neq" insertLatex="\neq " title="\neq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\approx" insertLatex="\approx " title="\approx" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\equiv" insertLatex="\equiv " title="\equiv" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\le" insertLatex="\le " title="\le" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ge" insertLatex="\ge " title="\ge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\forall" insertLatex="\forall " title="\forall" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\exists" insertLatex="\exists " title="\exists" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nexists" insertLatex="\nexists " title="\nexists" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\therefore" insertLatex="\therefore " title="\therefore" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\because" insertLatex="\because " title="\because" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\land" insertLatex="\land " title="\land" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lor" insertLatex="\lor " title="\lor" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\neg" insertLatex="\neg " title="\neg" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 3: TẬP HỢP */}
                <PivotItem headerText="Tập hợp">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\in" insertLatex="\in " title="\in" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\notin" insertLatex="\notin " title="\notin" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\subset" insertLatex="\subset " title="\subset" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\subseteq" insertLatex="\subseteq " title="\subseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\supset" insertLatex="\supset " title="\supset" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\cup" insertLatex="\cup " title="\cup" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\cap" insertLatex="\cap " title="\cap" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\setminus" insertLatex="\setminus " title="\setminus" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\emptyset" insertLatex="\emptyset " title="\emptyset" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\mathbb{R}" insertLatex="\mathbb{R}" title="\mathbb{R}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\mathbb{N}" insertLatex="\mathbb{N}" title="\mathbb{N}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\mathbb{Z}" insertLatex="\mathbb{Z}" title="\mathbb{Z}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    <MathBtn displayLatex="\mathbb{Q}" insertLatex="\mathbb{Q}" title="\mathbb{Q}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 4: GIẢI TÍCH */}
                <PivotItem headerText="Giải tích">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\int_{#0}^{#1}" insertLatex="\int_{#0}^{#1}" title="\int" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iint" insertLatex="\iint" title="\iint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oint" insertLatex="\oint" title="\oint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lim_{#? \to #?}" insertLatex="\lim_{#? \to #?}" title="\lim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sum" insertLatex="\sum_{#?}^{#?}" title="\sum" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\prod" insertLatex="\prod_{#?}^{#?}" title="\prod" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\infty" insertLatex="\infty" title="\infty" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\frac{d}{dx}" insertLatex="\frac{d}{dx} #0" title="\frac{d}{dx}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\partial" insertLatex="\partial " title="\partial" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nabla" insertLatex="\nabla " title="\nabla" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="!" insertLatex="! " title="!" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 5: HỆ & MA TRẬN */}
                <PivotItem headerText="Hệ & Ma trận">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
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
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
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

                        <MathBtn displayLatex="\hat{#?}" insertLatex="\hat{#0}" title="Góc trên đầu (ABC)" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\widehat{#?}" insertLatex="\widehat{#0}" title="Góc trên đầu (ABC)" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overset{\frown}{#0}" insertLatex="\overset{\frown}{#0}" title="Cung tròn" onInsert={onInsert} />
                    </Stack>
                </PivotItem>

                {/* TAB 7: LÝ & HÓA HỌC */}
                <PivotItem headerText="Lý & Hóa học">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow " title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightleftharpoons" insertLatex="\rightleftharpoons " title="\rightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uparrow" insertLatex="\uparrow " title="\uparrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downarrow" insertLatex="\downarrow " title="\downarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="t^\circ" insertLatex="t^\circ" title="t^\circ" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overset{t^\circ}{\rightarrow}" insertLatex="\overset{t^\circ}{\rightarrow}" title="\overset{t^\circ}{\rightarrow}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xrightarrow{\text{xt}}" insertLatex="\xrightarrow{\text{xt}} " title="\xrightarrow{\text{xt}}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="e^{-}" insertLatex="e^{-}" title="e^{-}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="^{A}_{Z}X}" insertLatex="^{#?}_{#?}#?}" title="^{A}_{Z}X}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Omega" insertLatex="\Omega " title="\Omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\omega" insertLatex="\omega " title="\omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lambda" insertLatex="\lambda " title="\lambda" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mu" insertLatex="\mu " title="\mu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{\AA}" insertLatex="\text{\AA}" title="\text{\AA}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hbar" insertLatex="\hbar " title="\hbar" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* Accents*/}
                <PivotItem headerText="Accents">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\bar{#?}" insertLatex="\bar{#0}" title="\bar" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\hat{#?}" insertLatex="\hat{#0}" title="\hat" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\widehat{#?}" insertLatex="\widehat{#0}" title="\widehat" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\tilde{#?}" insertLatex="\tilde{#0}" title="\tilde" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vec{#?}" insertLatex="\vec{#0}" title="\vec" onInsert={onInsert} />
                        <MathBtn displayLatex="\overline{#?}" insertLatex="\overline{#0}" title="\overline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underline{#?}" insertLatex="\underline{#0}" title="\underline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overrightarrow{#?}" insertLatex="\overrightarrow{#0}" title="\overrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underrightarrow{#?}" insertLatex="\underrightarrow{#0}" title="\underrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overbrace{#?}" insertLatex="\overbrace{#0}" title="\overbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underbrace{#?}" insertLatex="\underbrace{#0}" title="\underbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </Stack>
                </PivotItem>

                {/* TAB 8: MŨI TÊN */}
                <PivotItem headerText="Arrow">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\leftarrow" insertLatex="\leftarrow " title="\leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow " title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightarrow" insertLatex="\leftrightarrow " title="\leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftarrow" insertLatex="\Leftarrow " title="\Leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Rightarrow" insertLatex="\Rightarrow " title="\Rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftrightarrow" insertLatex="\Leftrightarrow " title="\Leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mapsto" insertLatex="\mapsto " title="\mapsto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nearrow" insertLatex="\nearrow " title="\nearrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\searrow" insertLatex="\searrow " title="\searrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\swarrow" insertLatex="\swarrow " title="\swarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nwarrow" insertLatex="\nwarrow " title="\nwarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xleftarrow{ab}" insertLatex="\xleftarrow{#0} " title="\xleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xLeftarrow{ab}" insertLatex="\xLeftarrow{#0} " title="\xLeftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xRightarrow{ab}" insertLatex="\xRightarrow{#0} " title="\xRightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightarrow{ab}" insertLatex="\xleftrightarrow{#0} " title="\xleftrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xLeftrightarrow{ab}" insertLatex="\xLeftrightarrow{#0} " title="\xLeftrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightharpoons{ab}" insertLatex="\xleftrightharpoons{#0} " title="\xleftrightharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightleftharpoons{ab}" insertLatex="\xrightleftharpoons{#0} " title="\xrightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xlongequal{ab}" insertLatex="\xlongequal{#0} " title="\xlongequal" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightarrow[xy]{ab}" insertLatex="xrightarrow[#0]{#?} " title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xmapsto{ab}" insertLatex="\xmapsto{#0} " title="\xmapsto" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </Stack>
                </PivotItem>

                {/* Annotation */}
                <PivotItem headerText="Annotation">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\cancel{2}" insertLatex="\cancel{#0}" title="\cancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bcancel{2}" insertLatex="\bcancel{#0}" title="\bcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xcancel{ABC}" insertLatex="\xcancel{ABC} " title="\xcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overbrace{a+b+c}^{\text{note}}" insertLatex="\overbrace{a+b+c}^{\text{#0}} " title="\overbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underbrace{a+b+c}_{\text{note}}" insertLatex="\underbrace{a+b+c}_{\text{#0}} " title="\underbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\boxed{AB}" insertLatex="\boxed{#0} " title="\boxed" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 9: HY LẠP */}
                <PivotItem headerText="Greek">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        {/* Chữ thường */}
                        <MathBtn displayLatex="\alpha" insertLatex="\alpha" title="\alpha" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\beta" insertLatex="\beta" title="\beta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\gamma" insertLatex="\gamma" title="\gamma" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\delta" insertLatex="\delta" title="\delta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\epsilon" insertLatex="\epsilon" title="\epsilon" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varepsilon" insertLatex="\varepsilon" title="\varepsilon" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\zeta" insertLatex="\zeta" title="\zeta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\eta" insertLatex="\eta" title="\eta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\theta" insertLatex="\theta" title="\theta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vartheta" insertLatex="\vartheta" title="\vartheta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iota" insertLatex="\iota" title="\iota" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\kappa" insertLatex="\kappa" title="\kappa" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lambda" insertLatex="\lambda" title="\lambda" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mu" insertLatex="\mu" title="\mu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nu" insertLatex="\nu" title="\nu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xi" insertLatex="\xi" title="\xi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\pi" insertLatex="\pi" title="\pi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varpi" insertLatex="\varpi" title="\varpi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rho" insertLatex="\rho" title="\rho" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varrho" insertLatex="\varrho" title="\varrho" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sigma" insertLatex="\sigma" title="\sigma" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varsigma" insertLatex="\varsigma" title="\varsigma" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\tau" insertLatex="\tau" title="\tau" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\upsilon" insertLatex="\upsilon" title="\upsilon" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\phi" insertLatex="\phi" title="\phi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varphi" insertLatex="\varphi" title="\varphi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\chi" insertLatex="\chi" title="\chi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\psi" insertLatex="\psi" title="\psi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\omega" insertLatex="\omega" title="\omega" onInsert={onInsert} borderColor={PaletteColors.default} />

                        {/* Chữ hoa */}
                        <MathBtn displayLatex="\Alpha" insertLatex="\Alpha" title="\Alpha" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Beta" insertLatex="\Beta" title="\Beta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Gamma" insertLatex="\Gamma" title="\Gamma" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Delta" insertLatex="\Delta" title="\Delta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Epsilon" insertLatex="\Epsilon" title="\Epsilon" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Zeta" insertLatex="\Zeta" title="\Zeta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Eta" insertLatex="\Eta" title="\Eta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Theta" insertLatex="\Theta" title="\Theta" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Iota" insertLatex="\Iota" title="\Iota" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Kappa" insertLatex="\Kappa" title="\Kappa" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Lambda" insertLatex="\Lambda" title="\Lambda" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Mu" insertLatex="\Mu" title="\Mu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Nu" insertLatex="\Nu" title="\Nu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Xi" insertLatex="\Xi" title="\Xi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Pi" insertLatex="\Pi" title="\Pi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Rho" insertLatex="\Rho" title="\Rho" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Sigma" insertLatex="\Sigma" title="\Sigma" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Tau" insertLatex="\Tau" title="\Tau" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Upsilon" insertLatex="\Upsilon" title="\Upsilon" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Phi" insertLatex="\Phi" title="\Phi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Chi" insertLatex="\Chi" title="\Chi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Psi" insertLatex="\Psi" title="\Psi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Omega" insertLatex="\Omega" title="\Omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

            </Pivot>
        </div>
    );
}