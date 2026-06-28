import * as React from"react";
import { DefaultButton, Stack, Pivot, PivotItem } from"@fluentui/react";

export const PaletteColors = {
    office:"#107c41", // 
    latex:"#f5b833", // Vàng chú ý
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
const MathBtn = ({ displayLatex, insertLatex, title, onInsert, borderColor }: { 
    displayLatex: string, 
    insertLatex: string, 
    title: string, 
    onInsert: (l: string) => void,
    borderColor?: string 
}) => {

    return (
        <DefaultButton
            title={title}
            onClick={() => onInsert(insertLatex)}
            onRenderDescription={() => null}
            styles={{
                root: {
                    minWidth:"40px",
                    height:"40px",
                    padding:"0",
                    position:"relative",
                    borderRadius:"5px",
                    backgroundColor:"transparent",
                    // Chỉ vẽ border nếu có borderColor
                    //border: borderColor ? `1px solid ${borderColor}` :"1px solid transparent"                   
                    border: borderColor ? `1.4px solid ${borderColor}` :"1px solid"
                },
                // Fluent UI tự động xử lý các trạng thái :hover, :active cho anh
                rootHovered: {
                    backgroundColor:"#e1dfdd", // Màu xám nhạt khi hover
                    borderColor: borderColor ? borderColor :"#a19f9d"
                },
                rootPressed: {
                    backgroundColor:"#c8c6c4" // Màu đậm hơn khi nhấn
                }
            }}
        >
            <div style={{ position:"absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 10, cursor:"pointer" }} />

            <math-field read-only style={{ pointerEvents:"none", border:"none", background:"transparent", padding: 0, margin: 0, minHeight:"auto", fontSize:"14px", fontWeight:"bold", outline:"none", color:"inherit" }}>
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
    const btnStyle = { minWidth:"42px", height:"42px", fontSize:"16px", padding:"0" };
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
        <div ref={pivotWrapperRef} style={{ position:"relative", marginBottom:"10px" }}>

            {/* NÚT BÊN TRÁI */}
            <DefaultButton
                onClick={() => scrollTabs('left')}
                style={{
                    position:"absolute", top: 0, left: 0, zIndex: 10,
                    minWidth:"32px", width:"32px", height:"44px", padding: 0,
                    border:"none", backgroundColor:"#faf9f8", color:"#107c41", fontWeight:"bold"
                }}
                title="Cuộn sang trái"
            >
                ❮
            </DefaultButton>

            {/* NÚT BÊN PHẢI */}
            <DefaultButton
                onClick={() => scrollTabs('right')}
                style={{
                    position:"absolute", top: 0, right: 0, zIndex: 10,
                    minWidth:"32px", width:"32px", height:"44px", padding: 0,
                    border:"none", backgroundColor:"#faf9f8", color:"#107c41", fontWeight:"bold"
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
                        overflowX:"auto", // Mở khóa tính năng cuộn nguyên bản (Vuốt Touchpad, Shift + Cuộn)
                        whiteSpace:"nowrap",
                        paddingLeft:"32px", // Chừa chỗ cho nút trái
                        paddingRight:"32px", // Chừa chỗ cho nút phải

                        // Tàng hình thanh trượt xám xịt
                        scrollbarWidth:"none",
                        msOverflowStyle:"none",
                        selectors: {
                           "&::-webkit-scrollbar": { display:"none" }
                        }
                    },
                    itemContainer: {
                        whiteSpace:"normal", // Đảm bảo các nút ở dưới không bị lỗi ngang
                        paddingTop:"5px"
                    },
                    link: { fontSize:"14px", fontWeight:"600", color:"#201f1e" },
                    linkIsSelected: { fontSize:"14px", fontWeight:"bold", color:"#107c41" }
                }}
            >

                {/* TAB 1: ĐẠI SỐ CƠ BẢN */}
                <PivotItem headerText="Basic">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\frac{#0}{#?}" insertLatex="\frac{#0}{#?}" title="\frac" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}/{#?}" insertLatex="{#0}/{#?}" title="/" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt{#0}" insertLatex="\sqrt{#0}" title="\sqrt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt[3]{#0}" insertLatex="\sqrt[3]{#0}" title="\sqrt[3]{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt[#?]{#0}" insertLatex="\sqrt[#?]{#0}" title="\sqrt[n]{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#?}^{#?}" insertLatex="{#0}^{#?}" title="^{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}_{#?}" insertLatex="{#0}_{#?}" title="_{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\left|#0\right|" insertLatex="\left|#0\right|" title="\left|\right|" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\overline{#0}" insertLatex="\overline{#0}" title="\overline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underline{#0}" insertLatex="\underline{#0}" title="\underline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </Stack>
                </PivotItem>

                {/* TAB 2: TOÁN TỬ & LOGIC */}
                <PivotItem headerText="Toán tử & Logic">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        
                        
                        <MathBtn displayLatex="\times" insertLatex="\times" title="\times" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\div" insertLatex="\div" title="\div" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdot" insertLatex="\cdot" title="\cdot" onInsert={onInsert} borderColor={PaletteColors.default} />
                         <MathBtn displayLatex="\pm" insertLatex="\pm" title="\pm" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mp" insertLatex="\mp" title="\mp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oplus" insertLatex="\oplus" title="\oplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ominus" insertLatex="\ominus" title="\ominus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\otimes" insertLatex="\otimes" title="\otimes" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\odot" insertLatex="\odot" title="\odot" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\approx" insertLatex="\approx" title="\approx" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\approxeq" insertLatex="\approxeq" title="\approxeq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\equiv" insertLatex="\equiv" title="\equiv" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\gt" insertLatex="\gt" title="\gt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ge" insertLatex="\ge" title="\ge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\geqslant" insertLatex="\geqslant" title="\geqslant" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lt" insertLatex="\lt" title="\tt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\le" insertLatex="\le" title="\le" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leqslant" insertLatex="\leqslant" title="\leqslant" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\propto" insertLatex="\propto" title="\propto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varpropto" insertLatex="\varpropto" title="\varpropto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangleq" insertLatex="\triangleq" title="\triangleq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\neq" insertLatex="\neq" title="\neq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\infty" insertLatex="\infty" title="\infty" onInsert={onInsert} borderColor={PaletteColors.default} />

                        
                    </Stack>
                </PivotItem>

                {/* TẬP HỢP */}
                <PivotItem headerText="Sets">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\forall" insertLatex="\forall" title="\forall" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\in" insertLatex="\in" title="\in" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\notin" insertLatex="\notin" title="\notin" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\not\in" insertLatex="\not\in" title="\not\in" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\subset" insertLatex="\subset" title="\subset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsubset" insertLatex="\nsubset" title="\nsubset" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\not\subset" insertLatex="\not\subset" title="\not\subset" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\subseteq" insertLatex="\subseteq" title="\subseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsubseteq" insertLatex="\nsubseteq" title="\nsubseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\supset" insertLatex="\supset" title="\supset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsupset" insertLatex="\nsupset" title="\nsupset" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\not\supset" insertLatex="\not\supset" title="\not\supset" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\cup" insertLatex="\cup" title="\cup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cap" insertLatex="\cap" title="\cap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\setminus" insertLatex="\setminus" title="\setminus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\emptyset" insertLatex="\emptyset" title="\emptyset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\forall" insertLatex="\forall" title="\forall" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\exists" insertLatex="\exists" title="\exists" onInsert={onInsert} borderColor={PaletteColors.default} />                     
                        <MathBtn displayLatex="\nexists" insertLatex="\nexists" title="\nexists" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\complement" insertLatex="\complement" title="\complement" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\mid" insertLatex="\mid" title="\mid" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\land" insertLatex="\land" title="\land" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lor" insertLatex="\lor" title="\lor" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ni" insertLatex="\ni" title="\ni" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\therefore" insertLatex="\therefore" title="\therefore" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\because" insertLatex="\because" title="\because" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mapsto" insertLatex="\mapsto" title="\mapsto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\to" insertLatex="\to" title="\to" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\gets" insertLatex="\gets" title="\gets" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightarrow" insertLatex="\leftrightarrow" title="\leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\emptyset" insertLatex="\emptyset" title="\emptyset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\empty" insertLatex="\empty" title="\empty" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varnothing" insertLatex="\varnothing" title="\varnothing" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\implies" insertLatex="\implies" title="\implies" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\impliedby" insertLatex="\impliedby" title="\impliedby" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iff" insertLatex="\iff" title="\iff" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\neg" insertLatex="\neg" title="\neg" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lnot" insertLatex="\lnot" title="\lnot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\mathbb{N}" insertLatex="\mathbb{N}" title="\mathbb{N}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mathbb{Z}" insertLatex="\mathbb{Z}" title="\mathbb{Z}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mathbb{R}" insertLatex="\mathbb{R}" title="\mathbb{R}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mathbb{Q}" insertLatex="\mathbb{Q}" title="\mathbb{Q}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* GIẢI TÍCH */}
                <PivotItem headerText="Giải tích">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\int_{#0}^{#1}" insertLatex="\int_{#0}^{#1}" title="\int" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iint" insertLatex="\iint" title="\iint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oint" insertLatex="\oint" title="\oint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lim_{#?}" insertLatex="\lim_{#?}" title="\lim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lim_{#? \to #?}" insertLatex="\lim_{#? \to #?}" title="\lim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\log_{#0}{}" insertLatex="\log_{#0}{#?}" title="\log_{}{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sum" insertLatex="\sum_{#?}^{#?}" title="\sum" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\prod" insertLatex="\prod_{#?}^{#?}" title="\prod" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\coprod" insertLatex="\coprod_{#?}^{#?}" title="\coprod" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\infty" insertLatex="\infty" title="\infty" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\frac{d}{dx}" insertLatex="\frac{d}{dx} #0" title="\frac{d}{dx}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\partial" insertLatex="\partial" title="\partial" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nabla" insertLatex="\nabla" title="\nabla" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="!" insertLatex="!" title="!" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* TAB 5: HỆ & MA TRẬN */}
                <PivotItem headerText="Hệ & Ma trận">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\lbrace\text{Hệ 2}" insertLatex="\begin{cases} #? &= #? \\ #? &= #? \end{cases}" title="\begin{cases} (2pt)" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lbrace\text{Hệ 3}" insertLatex="\begin{cases} #? &= #? \\ #? &= #? \\ #? &= #? \end{cases}" title="\begin{cases} (3pt)" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\left[ \text{Hoặc}" insertLatex="\left[ \begin{array}{l} #? \\ #? \end{array} \right." title="\left[ \begin{array}{l}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="f(x) = \lbrace" insertLatex="f(x) = \begin{cases} #?, & \text{nếu } #? \\ #?, & \text{nếu } #? \end{cases}" title="f(x) \begin{cases}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{(■)}" insertLatex="\begin{pmatrix} #? & #? \\ #? & #? \end{pmatrix}" title="\begin{pmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{[■]}" insertLatex="\begin{bmatrix} #? & #? \\ #? & #? \end{bmatrix}" title="\begin{bmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{|■|}" insertLatex="\begin{vmatrix} #? & #? \\ #? & #? \end{vmatrix}" title="\begin{vmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{[⋮]}" insertLatex="\begin{bmatrix} #? \\ #? \\ #? \end{bmatrix}" title="\begin{bmatrix} 3x1" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdots" insertLatex="\cdots" title="\cdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vdots" insertLatex="\vdots" title="\vdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ddots" insertLatex="\ddots" title="\ddots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                    </Stack>
                </PivotItem>

                {/* HÌNH & LƯỢNG GIÁC */}
                <PivotItem headerText="Hình & Lượng giác">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\sin" insertLatex="\sin(#?)" title="\sin" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sin^{-1}" insertLatex="\sin^{-1}(#?)" title="\sin^{-1}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cos" insertLatex="\cos(#?)" title="\cos" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cos^{-1}" insertLatex="\cos^{-1}(#?)" title="\cos^{-1}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\tan" insertLatex="\tan(#?)" title="\tan" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\tan^{-1}" insertLatex="\tan^{-1}(#?)" title="\tan^{-1}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cot" insertLatex="\cot(#?)" title="\cot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cot^{-1}" insertLatex="\cot^{-1}(#?)" title="\cot^{-1}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\pi" insertLatex="\pi" title="\pi" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="^\circ" insertLatex="^{\circ}" title="^{\circ}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangle" insertLatex="\triangle" title="\triangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\angle" insertLatex="\angle" title="\angle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\perp" insertLatex="\perp" title="\perp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\parallel" insertLatex="\parallel" title="\parallel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="ᔕ" insertLatex="\text{ᔕ}" title="\text{ᔕ}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\backsim" insertLatex="\backsim" title="\backsim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vec{#?}" insertLatex="\vec{#0}" title="Vector" onInsert={onInsert} />

                        <MathBtn displayLatex="\hat{#?}" insertLatex="\hat{#0}" title="\hat{}" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\widehat{#?}" insertLatex="\widehat{#0}" title="\widehat{}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overset{\frown}{#0}" insertLatex="\overset{\frown}{#0}" title="\overset{\frown}{}" onInsert={onInsert} />
                        <MathBtn displayLatex="\overset{\curvearrowright}{#0}" insertLatex="\overset{\curvearrowright}{#0}" title="\overset{\curvearrowright}{}" onInsert={onInsert} />
                    </Stack>
                </PivotItem>

                {/* LÝ & HÓA HỌC */}
                <PivotItem headerText="Lý & Hóa học">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow" title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightleftharpoons" insertLatex="\rightleftharpoons" title="\rightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uparrow" insertLatex="\uparrow" title="\uparrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downarrow" insertLatex="\downarrow" title="\downarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="t^\circ" insertLatex="t^\circ" title="t^\circ" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overset{t^\circ}{\rightarrow}" insertLatex="\overset{t^\circ}{\rightarrow}" title="\overset{t^\circ}{\rightarrow}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xrightarrow{\text{xt}}" insertLatex="\xrightarrow{\text{xt}}" title="\xrightarrow{\text{xt}}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="e^{-}" insertLatex="e^{-}" title="e^{-}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="^{A}_{Z}X}" insertLatex="^{#?}_{#?}#?}" title="^{A}_{Z}X}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\Omega" insertLatex="\Omega" title="\Omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\omega" insertLatex="\omega" title="\omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lambda" insertLatex="\lambda" title="\lambda" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mu" insertLatex="\mu" title="\mu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{\AA}" insertLatex="\text{\AA}" title="\text{\AA}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hbar" insertLatex="\hbar" title="\hbar" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* Accents*/}
                <PivotItem headerText="Accents">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\bar{#?}" insertLatex="\bar{#0}" title="\bar" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\bar{\bar{a}}" insertLatex="\bar{\bar{#0}}" title="\bar{\bar{}}" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\hat{#?}" insertLatex="\hat{#0}" title="\hat" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\widehat{#?}" insertLatex="\widehat{#0}" title="\widehat" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overline{#?}" insertLatex="\overline{#0}" title="\overline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overline{\overline{#?}}" insertLatex="\overline{\overline{#0}}" title="\overline{\overline{}}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underline{#?}" insertLatex="\underline{#0}" title="\underline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\tilde{#?}" insertLatex="\tilde{#0}" title="\tilde" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vec{#?}" insertLatex="\vec{#0}" title="\vec" onInsert={onInsert} />
                        
                        <MathBtn displayLatex="\overrightarrow{#?}" insertLatex="\overrightarrow{#0}" title="\overrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underrightarrow{#?}" insertLatex="\underrightarrow{#0}" title="\underrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overbrace{#?}" insertLatex="\overbrace{#0}" title="\overbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underbrace{#?}" insertLatex="\underbrace{#0}" title="\underbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </Stack>
                </PivotItem>

                {/* MŨI TÊN */}
                <PivotItem headerText="Arrows">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\leftarrow" insertLatex="\leftarrow" title="\leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow" title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightarrow" insertLatex="\leftrightarrow" title="\leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftarrow" insertLatex="\Leftarrow" title="\Leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Rightarrow" insertLatex="\Rightarrow" title="\Rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftrightarrow" insertLatex="\Leftrightarrow" title="\Leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mapsto" insertLatex="\mapsto" title="\mapsto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nearrow" insertLatex="\nearrow" title="\nearrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\searrow" insertLatex="\searrow" title="\searrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\swarrow" insertLatex="\swarrow" title="\swarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nwarrow" insertLatex="\nwarrow" title="\nwarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xrightarrow{ab}" insertLatex="\xrightarrow{#0}" title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\overset{ab}{\rightarrow}" insertLatex="\overset{#0}{\rightarrow}" title="\overset{}{\rightarrow}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xleftarrow{ab}" insertLatex="\xleftarrow{#0}" title="\xleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xLeftarrow{ab}" insertLatex="\xLeftarrow{#0}" title="\xLeftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xRightarrow{ab}" insertLatex="\xRightarrow{#0}" title="\xRightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightarrow{ab}" insertLatex="\xleftrightarrow{#0}" title="\xleftrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xLeftrightarrow{ab}" insertLatex="\xLeftrightarrow{#0}" title="\xLeftrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightharpoons{ab}" insertLatex="\xleftrightharpoons{#0}" title="\xleftrightharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightleftharpoons{ab}" insertLatex="\xrightleftharpoons{#0}" title="\xrightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xlongequal{ab}" insertLatex="\xlongequal{#0}" title="\xlongequal" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightarrow[xy]{ab}" insertLatex="\xrightarrow[#0]{#?}" title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xmapsto{ab}" insertLatex="\xmapsto{#0}" title="\xmapsto" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </Stack>
                </PivotItem>

                {/* Symbols */}
                <PivotItem headerText="Symbols">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\sim" insertLatex="\sim" title="\sim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\backsim" insertLatex="\backsim" title="\backsim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\infty" insertLatex="\infty" title="\infty" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nabla" insertLatex="\nabla" title="\nabla" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\partial" insertLatex="\partial" title="\partial" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\eth" insertLatex="\eth" title="\eth" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\clubsuit" insertLatex="\clubsuit" title="\clubsuit" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\diamondsuit" insertLatex="\diamondsuit" title="\diamondsuit" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\heartsuit" insertLatex="\heartsuit" title="\heartsuit" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\spadesuit" insertLatex="\spadesuit" title="\spadesuit" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdots" insertLatex="\cdots" title="\cdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vdots" insertLatex="\vdots" title="\vdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ldots" insertLatex="\ldots" title="\ldots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ddots" insertLatex="\ddots" title="\ddots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Im" insertLatex="\Im" title="\Im" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Re" insertLatex="\Re" title="\Re" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="ᔕ" insertLatex="\text{ᔕ}" title="\text{ᔕ}" onInsert={onInsert} borderColor={PaletteColors.default} />

                       
                        <MathBtn displayLatex="\imath" insertLatex="\imath" title="\imath" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\jmath" insertLatex="\jmath" title="\jmath" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ell" insertLatex="\ell" title="\ell" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iiint" insertLatex="\iiint" title="\iiint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\iint" insertLatex="\iint" title="\iint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sharp" insertLatex="\sharp" title="\sharp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\flat" insertLatex="\flat" title="\flat" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\natural" insertLatex="\natural" title="\natural" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\Bbbk" insertLatex="\Bbbk" title="\Bbbk" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigstar" insertLatex="\bigstar" title="\bigstar" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\diagup" insertLatex="\diagup" title="\diagup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Diamond" insertLatex="\Diamond" title="\Diamond" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Finv" insertLatex="\Finv" title="\Finv" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Game" insertLatex="\Game" title="\Game" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hbar" insertLatex="\hbar" title="\hbar" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hslash" insertLatex="\hslash" title="\hslash" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lozenge" insertLatex="\lozenge" title="\lozenge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mho" insertLatex="\mho" title="\mho" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\prime" insertLatex="\prime" title="\prime" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\square" insertLatex="\square" title="\square" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\surd" insertLatex="\surd" title="\surd" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\wp" insertLatex="\wp" title="\wp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\angle" insertLatex="\angle" title="\angle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\measuredangle" insertLatex="\measuredangle" title="\measuredangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sphericalangle" insertLatex="\sphericalangle" title="\sphericalangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\complement" insertLatex="\complement" title="\complement" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangledown" insertLatex="\triangledown" title="\triangledown" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangle" insertLatex="\triangle" title="\triangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vartriangle" insertLatex="\vartriangle" title="\vartriangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\blacklozenge" insertLatex="\blacklozenge" title="\blacklozenge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\blacksquare" insertLatex="\blacksquare" title="\blacksquare" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\blacktriangle" insertLatex="\blacktriangle" title="\blacktriangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\blacktriangledown" insertLatex="\blacktriangledown" title="\blacktriangledown" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\backprime" insertLatex="\backprime" title="\backprime" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\circledS" insertLatex="\circledS" title="\circledS" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{♂}" insertLatex="\text{♂}" title="\text{♂}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{♀}" insertLatex="\text{♀}" title="\text{♀}" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </Stack>
                </PivotItem>

                {/* Annotation */}
                <PivotItem headerText="Annotation">
                    <Stack horizontal wrap tokens={{ childrenGap: 2 }} style={{ marginTop: 5 }}>
                        <MathBtn displayLatex="\cancel{2}" insertLatex="\cancel{#0}" title="\cancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bcancel{2}" insertLatex="\bcancel{#0}" title="\bcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xcancel{ABC}" insertLatex="\xcancel{ABC}" title="\xcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overbrace{a+b+c}^{\text{note}}" insertLatex="\overbrace{a+b+c}^{\text{#0}}" title="\overbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underbrace{a+b+c}_{\text{note}}" insertLatex="\underbrace{a+b+c}_{\text{#0}}" title="\underbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\boxed{AB}" insertLatex="\boxed{#0}" title="\boxed" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\phase{60^\circ}" insertLatex="\phase{#0}" title="\phase{}" onInsert={onInsert} borderColor={PaletteColors.latex} />
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