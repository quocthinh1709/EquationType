import * as React from "react";
import { useState, useRef } from "react";
import { DefaultButton, Stack, Pivot, PivotItem, Callout, DirectionalHint } from "@fluentui/react";

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

// Hàm tạo nút bấm chứa công thức hiển thị trực quan (Dành cho các nút tĩnh)
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
                    minWidth: "40px",
                    height: "40px",
                    padding: "0",
                    position: "relative",
                    borderRadius: "5px",
                    backgroundColor: "transparent",
                    //border: borderColor ? `1.4px solid ${borderColor}` : "1px solid"
                    borderColor: "transparent"
                },
                rootHovered: {
                    backgroundColor: "#e1dfdd",
                    //borderColor: borderColor ? borderColor : "#a19f9d"
                    borderColor: borderColor ? borderColor : "${borderColor}"
                },
                rootPressed: {
                    backgroundColor: "#c8c6c4"
                }
            }}
        >
            <div style={{ position: "absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 10, cursor: "pointer" }} />
            <math-field read-only style={{ pointerEvents: "none", border: "none", background: "transparent", padding: 0, margin: 0, minHeight: "auto", fontSize: "18px", outline: "none", color: "inherit" }}>
                {displayLatex}
            </math-field>
        </DefaultButton>
    );
};

// ==========================================
// THÊM MỚI: Thanh dọc phân cách các nhóm nút
// ==========================================
const VerticalDivider = () => {
    return (
        <div style={{
            width: "1px",
            height: "32px",               // Chiều cao ngắn hơn nút một chút (nút của anh cao 40px)
            backgroundColor: "#c8c6c4",   // Màu xám nhạt chuẩn của giao diện Office
            margin: "4px 4px",            // Khoảng cách hai bên trái phải
            alignSelf: "center"           // Căn giữa theo chiều dọc
        }} />
    );
};

// ==========================================
// THÊM MỚI: BỘ GOM NÚT TỰ ĐỘNG BẰNG CÁCH ĐO TỌA ĐỘ THỰC TẾ
// ==========================================
const OverflowMathGroup = ({ children }: { children: React.ReactNode }) => {
    const [isCalloutVisible, setIsCalloutVisible] = useState(false);

    // maxVisible = null nghĩa là đang trong trạng thái "Đo lường"
    const [maxVisible, setMaxVisible] = useState<number | null>(null);

    const containerRef = useRef<HTMLDivElement>(null);
    const buttonRef = useRef<HTMLDivElement>(null);

    const childrenArray = React.Children.toArray(children);

    // 1. Hàm thực hiện đo lường vị trí các nút
    React.useEffect(() => {
        // Chỉ đo lường khi maxVisible bị reset về null
        if (maxVisible !== null) return undefined;

        const calculateCapacity = () => {
            if (!containerRef.current) return;

            // Lấy thẻ div chứa toàn bộ các nút
            const flexContainer = containerRef.current.firstElementChild as HTMLElement;
            if (!flexContainer) return;

            const items = Array.from(flexContainer.children) as HTMLElement[];
            if (items.length === 0) return;

            // Lấy tọa độ Y (offsetTop) của nút đầu tiên làm mốc của Hàng 1
            const firstItemTop = items[0].offsetTop;
            let currentRow = 1;
            let lastItemTop = firstItemTop;
            let splitIndex = items.length;

            // Quét qua từng nút để xem nó nằm ở hàng thứ mấy
            for (let i = 0; i < items.length; i++) {
                const item = items[i];

                // Nếu tọa độ Y của nút này lớn hơn nút trước đó 20px -> Nó đã rớt xuống hàng mới
                if (item.offsetTop > lastItemTop + 20) {
                    currentRow++;
                    lastItemTop = item.offsetTop;
                }

                // Ngay khi phát hiện có nút rớt xuống HÀNG 3
                if (currentRow > 2) {
                    splitIndex = i; // Đánh dấu vị trí bắt đầu bị tràn
                    break;
                }
            }

            // Nếu có tràn hàng 3 -> Cắt bớt 1 nút cuối cùng của hàng 2 để nhường chỗ cho Nút Mũi Tên
            if (splitIndex < items.length) {
                setMaxVisible(splitIndex - 1);
            } else {
                setMaxVisible(items.length); // Đủ chỗ thì hiện hết
            }
        };

        // Chờ 100ms để Font chữ MathLive render xong hình khối rồi mới đo
        const timer = setTimeout(calculateCapacity, 100);
        return () => clearTimeout(timer);
    }, [maxVisible, childrenArray.length]);

    // 2. Lắng nghe sự kiện người dùng kéo giãn/thu hẹp khung Word
    React.useEffect(() => {
        let resizeTimer: any;
        const handleResize = () => {
            clearTimeout(resizeTimer);
            // Khi bị kéo khung, ép hệ thống đo lại từ đầu
            resizeTimer = setTimeout(() => setMaxVisible(null), 150);
        };
        window.addEventListener("resize", handleResize);
        return () => window.removeEventListener("resize", handleResize);
    }, []);


    // === PHẦN RENDER GIAO DIỆN ===
    const isMeasuring = maxVisible === null;

    // Khi đang đo: Render TẤT CẢ nút. Khi đo xong: Chỉ render số nút vừa vặn 2 hàng
    const visibleChildren = isMeasuring ? childrenArray : childrenArray.slice(0, maxVisible);
    const hiddenChildren = isMeasuring ? [] : childrenArray.slice(maxVisible);

    return (
        <div
            ref={containerRef}
            style={{
                width: "100%",
                // TRICK: Khi đang đo, ép chiều cao khung tối đa 95px và giấu đi phần tràn. 
                // Nhờ vậy, người dùng sẽ không thấy giao diện bị chớp giật hay lòi ra hàng 3!
                maxHeight: isMeasuring ? "95px" : "auto",
                overflow: isMeasuring ? "hidden" : "visible"
            }}
        >
            {/* Sử dụng thẻ div flex thuần thay vì Stack để đo tọa độ chính xác tuyệt đối */}
            <div style={{ display: "flex", flexWrap: "wrap", gap: "2px", marginTop: "5px" }}>

                {/* 1. Các nút hiển thị chính */}
                {visibleChildren}

                {/* 2. Nút mũi tên (chỉ hiện khi đã đo xong và phát hiện có nút bị giấu) */}
                {!isMeasuring && hiddenChildren.length > 0 && (
                    <div ref={buttonRef} style={{ display: "inline-block" }}>
                        <DefaultButton
                            title="Xem thêm ký hiệu"
                            onClick={() => setIsCalloutVisible(!isCalloutVisible)}
                            iconProps={{ iconName: 'ChevronDown' }}
                            styles={{
                                root: {
                                    minWidth: "40px",
                                    height: "40px",
                                    padding: "0",
                                    borderRadius: "5px",
                                    backgroundColor: isCalloutVisible ? "#e1dfdd" : "transparent",
                                    borderColor: "transparent"
                                },
                                rootHovered: { backgroundColor: "#e1dfdd" },
                                rootPressed: { backgroundColor: "#c8c6c4" },
                                icon: { fontSize: 16, color: "#107c41", fontWeight: "bold" }
                            }}
                        />

                        {isCalloutVisible && (
                            <Callout
                                target={buttonRef}
                                onDismiss={() => setIsCalloutVisible(false)}
                                directionalHint={DirectionalHint.bottomRightEdge}
                                isBeakVisible={true}
                                styles={{ calloutMain: { padding: "5px", backgroundColor: "#ffffff", maxWidth: "100px" } }}
                            >
                                {/* Đã thay thế onClickCapture thành onClick để fix lỗi không chèn được chữ */}
                                <div onClick={() => setIsCalloutVisible(false)}>
                                    <div style={{ display: "flex", flexWrap: "wrap", gap: "2px" }}>
                                        {hiddenChildren}
                                    </div>
                                </div>
                            </Callout>
                        )}
                    </div>
                )}
            </div>
        </div>
    );
};
// ==========================================
// THÊM MỚI: HÀM SINH LATEX MA TRẬN ĐỘNG
// ==========================================
const generateMatrixLatex = (matrixType: string, rows: number, cols: number) => {
    let latex = `\\begin{${matrixType}}\n`;
    for (let r = 0; r < rows; r++) {
        let rowContent = [];
        for (let c = 0; c < cols; c++) {
            rowContent.push("#?"); // Placeholder cho MathLive
        }
        latex += rowContent.join(" & ");
        if (r < rows - 1) latex += " \\\\\n";
    }
    latex += `\n\\end{${matrixType}}`;
    return latex;
};

// =========================================================================
// LOGIC: Hàm sinh mã LaTeX động cho Hệ Phương Trình & Hàm Số (Đã làm tổng quát)
// =========================================================================
const generateSystemLatex = (type: "cases" | "array" | "matrix" | "piecewise", rows: number, cols: number) => {
    let latex = "";

    if (type === "cases") {
        // Hệ và (Ngoặc nhọn lớn)
        latex += "\\begin{cases}\n";
        for (let r = 0; r < rows; r++) {
            let rowContent = Array(cols).fill("#?");
            latex += rowContent.join(" & ");
            if (r < rows - 1) latex += " \\\\\n";
        }
        latex += "\n\\end{cases}";

    } else if (type === "array") {
        // Hệ hoặc (Ngoặc vuông bên trái)
        const alignStr = Array(cols).fill("l").join("");
        latex += `\\left[ \\begin{array}{${alignStr}}\n`;
        for (let r = 0; r < rows; r++) {
            let rowContent = Array(cols).fill("#?");
            latex += rowContent.join(" & ");
            if (r < rows - 1) latex += " \\\\\n";
        }
        latex += "\n\\end{array} \\right.";

    } else if (type === "matrix") {
        // Hệ hoặc (Ngoặc vuông bên trái)

        latex += `\\left[ \\begin{matrix}\n`;
        for (let r = 0; r < rows; r++) {
            let rowContent = Array(cols).fill("#?");
            latex += rowContent.join(" & ");
            if (r < rows - 1) latex += " \\\\\n";
        }
        latex += "\n\\end{matrix} \\right.";

    } else if (type === "piecewise") {
        // Hàm số phân nhánh f(x) = { (Thuần túy, không có chữ "nếu")
        latex += "f(x) = \\begin{cases}\n";
        for (let r = 0; r < rows; r++) {
            let rowContent = Array(cols).fill("#?");
            latex += rowContent.join(" & ");
            if (r < rows - 1) latex += " \\\\\n";
        }
        latex += "\n\\end{cases}";
    }

    return latex;
};

// =========================================================================
// COMPONENT: Nút chọn Hệ phương trình / Hàm số (Lưới 2D: Dòng x Cột)
// =========================================================================
const SystemSelectorBtn = ({ displayLatex, type, title, onInsert, borderColor }: {
    displayLatex: string,
    type: "cases" | "array" | "matrix" | "piecewise",
    title: string,
    onInsert: (l: string) => void,
    borderColor?: string
}) => {
    const [isCalloutVisible, setIsCalloutVisible] = useState(false);
    const [hoveredSize, setHoveredSize] = useState({ r: 0, c: 0 });
    const buttonRef = useRef<HTMLDivElement>(null);
    const maxRows = 6;
    const maxCols = 6; // Với hệ phương trình, 5 cột (VD: a, x, b, y, c) là khá dư dả

    const handleSelect = (r: number, c: number) => {
        const latex = generateSystemLatex(type, r, c);
        onInsert(latex);
        setIsCalloutVisible(false);
        setHoveredSize({ r: 0, c: 0 });
    };

    return (
        <div ref={buttonRef} style={{ display: "inline-block" }}>
            <DefaultButton
                title={title}
                onClick={() => setIsCalloutVisible(!isCalloutVisible)}
                styles={{
                    root: {
                        minWidth: "40px",
                        height: "40px",
                        padding: "0",
                        position: "relative",
                        borderRadius: "5px",
                        backgroundColor: isCalloutVisible ? "#e1dfdd" : "transparent",
                        //border: borderColor ? `1.4px solid ${borderColor}` : "1px solid",
                        borderColor: "transparent"
                    },
                    rootHovered: {
                        backgroundColor: "#e1dfdd",
                        //borderColor: borderColor ? borderColor : "#a19f9d" 
                        borderColor: borderColor ? borderColor : "${borderColor}"
                    },
                    rootPressed: { backgroundColor: "#c8c6c4" }
                }}
            >
                <div style={{ position: "absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 10, cursor: "pointer" }} />
                <math-field read-only style={{ pointerEvents: "none", border: "none", background: "transparent", padding: 0, margin: 0, minHeight: "auto", fontSize: "14px", fontWeight: "bold", outline: "none", color: "inherit" }}>
                    {displayLatex}
                </math-field>
            </DefaultButton>

            {isCalloutVisible && (
                <Callout
                    target={buttonRef}
                    onDismiss={() => setIsCalloutVisible(false)}
                    directionalHint={DirectionalHint.bottomCenter}
                    isBeakVisible={true}
                    styles={{ calloutMain: { padding: "10px", backgroundColor: "#ffffff" } }}
                >
                    <div style={{ textAlign: "center", marginBottom: "8px", fontWeight: "bold", fontSize: "12px", color: "#107c41" }}>
                        {hoveredSize.r > 0 ? `Size: ${hoveredSize.r} rows x ${hoveredSize.c} columns` : "Choose size"}
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: "2px" }} onMouseLeave={() => setHoveredSize({ r: 0, c: 0 })}>
                        {Array.from({ length: maxRows }).map((_, rowIndex) => (
                            <div key={rowIndex} style={{ display: "flex", gap: "2px" }}>
                                {Array.from({ length: maxCols }).map((_, colIndex) => {
                                    const r = rowIndex + 1; // Bắt đầu từ 1 dòng
                                    const c = colIndex + 1; // Bắt đầu từ 1 cột
                                    const isHovered = r <= hoveredSize.r && c <= hoveredSize.c && hoveredSize.r > 0;

                                    return (
                                        <div
                                            key={colIndex}
                                            onMouseEnter={() => setHoveredSize({ r, c })}
                                            onClick={() => handleSelect(r, c)}
                                            style={{
                                                width: "20px",
                                                height: "20px",
                                                border: "1px solid #c8c6c4",
                                                backgroundColor: isHovered ? "#b4d6fa" : "#ffffff",
                                                cursor: "pointer",
                                                borderRadius: "2px"
                                            }}
                                        />
                                    );
                                })}
                            </div>
                        ))}
                    </div>
                </Callout>
            )}
        </div>
    );
};
// ==========================================
// THÊM MỚI: NÚT BẤM CÓ HIỂN THỊ POPUP CHỌN KÍCH THƯỚC MA TRẬN
// ==========================================
const MatrixSelectorBtn = ({ displayLatex, matrixType, title, onInsert, borderColor }: {
    displayLatex: string,
    matrixType: string, // VD: pmatrix, bmatrix, vmatrix
    title: string,
    onInsert: (l: string) => void,
    borderColor?: string
}) => {
    const [isCalloutVisible, setIsCalloutVisible] = useState(false);
    const [hoveredSize, setHoveredSize] = useState({ r: 0, c: 0 });
    const buttonRef = useRef<HTMLDivElement>(null);
    const maxRows = 6;
    const maxCols = 6;

    const handleSelect = (r: number, c: number) => {
        const latex = generateMatrixLatex(matrixType, r, c);
        onInsert(latex);
        setIsCalloutVisible(false);
        setHoveredSize({ r: 0, c: 0 }); // Reset
    };

    return (
        <div ref={buttonRef} style={{ display: "inline-block" }}>
            <DefaultButton
                title={title}
                onClick={() => setIsCalloutVisible(!isCalloutVisible)}
                onRenderDescription={() => null}
                styles={{
                    root: {
                        minWidth: "40px",
                        height: "40px",
                        padding: "0",
                        position: "relative",
                        borderRadius: "5px",
                        backgroundColor: isCalloutVisible ? "#e1dfdd" : "transparent",
                        //border: borderColor ? `1.4px solid ${borderColor}` : "1px solid"
                        borderColor: "transparent"
                    },
                    rootHovered: {
                        backgroundColor: "#e1dfdd",
                        //borderColor: borderColor ? borderColor : "#a19f9d" 
                        borderColor: borderColor ? borderColor : "${borderColor}"
                    },
                    rootPressed: { backgroundColor: "#c8c6c4" }
                }}
            >
                <div style={{ position: "absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 10, cursor: "pointer" }} />
                <math-field read-only style={{ pointerEvents: "none", border: "none", background: "transparent", padding: 0, margin: 0, minHeight: "auto", fontSize: "14px", fontWeight: "bold", outline: "none", color: "inherit" }}>
                    {displayLatex}
                </math-field>
            </DefaultButton>

            {isCalloutVisible && (
                <Callout
                    target={buttonRef}
                    onDismiss={() => setIsCalloutVisible(false)}
                    directionalHint={DirectionalHint.bottomCenter}
                    isBeakVisible={true}
                    styles={{ calloutMain: { padding: "10px", backgroundColor: "#ffffff" } }}
                >
                    <div style={{ textAlign: "center", marginBottom: "8px", fontWeight: "bold", fontSize: "12px", color: "#107c41" }}>
                        {hoveredSize.r > 0 ? `${hoveredSize.r} x ${hoveredSize.c} Matrix` : "Chọn kích thước"}
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: "2px" }} onMouseLeave={() => setHoveredSize({ r: 0, c: 0 })}>
                        {Array.from({ length: maxRows }).map((_, rowIndex) => (
                            <div key={rowIndex} style={{ display: "flex", gap: "2px" }}>
                                {Array.from({ length: maxCols }).map((_, colIndex) => {
                                    const r = rowIndex + 1;
                                    const c = colIndex + 1;
                                    const isHovered = r <= hoveredSize.r && c <= hoveredSize.c;
                                    return (
                                        <div
                                            key={colIndex}
                                            onMouseEnter={() => setHoveredSize({ r, c })}
                                            onClick={() => handleSelect(r, c)}
                                            style={{
                                                width: "20px",
                                                height: "20px",
                                                border: "1px solid #c8c6c4",
                                                backgroundColor: isHovered ? "#b4d6fa" : "#ffffff", // Màu xanh dương nhạt khi di chuột
                                                cursor: "pointer",
                                                borderRadius: "2px"
                                            }}
                                        />
                                    );
                                })}
                            </div>
                        ))}
                    </div>
                </Callout>
            )}
        </div>
    );
};

// ==========================================
// THÊM MỚI: Hàm hỗ trợ hiển thị LaTeX trên tiêu đề Tab (Đã fix lỗi click)
// ==========================================
const renderLatexTab = (latexStr: string) => {
    return (
        <span style={{
            position: "relative", // Để làm mỏ neo cho lớp phủ absolute bên trong
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            width: "100%",
            height: "100%",
        }}>

            <div style={{ position: "absolute", top: 0, left: 0, right: 0, bottom: 0, zIndex: 8, cursor: "pointer" }} />

            <math-field
                read-only
                tabIndex={-1}
                style={{
                    pointerEvents: "none",
                    userSelect: "none",
                    border: "none",
                    background: "transparent",
                    padding: 0,
                    margin: 0,
                    minHeight: "auto",
                    fontSize: "14px",
                    fontWeight: "bold",
                    outline: "none",
                    color: "inherit"
                }}
            >
                {latexStr}
            </math-field>
        </span>
    );
};

export interface MathPaletteProps {
    onInsert: (latexCode: string) => void;
}

export default function MathPalette({ onInsert }: MathPaletteProps) {
    const pivotWrapperRef = React.useRef<HTMLDivElement>(null);
    const scrollTabs = (direction: 'left' | 'right') => {
        if (pivotWrapperRef.current) {
            // Trở về cách gọi gốc: Tìm đúng thẻ tablist của Fluent UI
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
        <div ref={pivotWrapperRef} style={{ position: "relative", marginBottom: "10px" }}>
            <DefaultButton
                onClick={() => scrollTabs('left')}
                style={{ position: "absolute", top: 0, left: 0, zIndex: 50, minWidth: "32px", width: "32px", height: "44px", padding: 0, border: "none", backgroundColor: "#faf9f8", color: "#107c41", fontWeight: "bold" }}
            >❮</DefaultButton>

            <DefaultButton
                onClick={() => scrollTabs('right')}
                style={{ position: "absolute", top: 0, right: 0, zIndex: 50, minWidth: "32px", width: "32px", height: "44px", padding: 0, border: "none", backgroundColor: "#faf9f8", color: "#107c41", fontWeight: "bold" }}
            >❯</DefaultButton>

            <Pivot aria-label="Math Palette Full"
                styles={{
                    root: {
                        overflowX: "auto",
                        whiteSpace: "nowrap",
                        paddingLeft: "34px",
                        paddingRight: "34px",
                        scrollbarWidth: "none",
                        msOverflowStyle: "none",
                        borderRadius: "15px",
                        selectors: { "&::-webkit-scrollbar": { display: "none" } }
                    },
                    itemContainer: {
                        whiteSpace: "normal",
                        paddingTop: "5px"
                    },
                    link: {
                        fontSize: "14px",
                        fontWeight: "600",
                        color: "#201f1e",
                        backgroundColor: "#e1e1e1",
                        width: "50px",
                        minWidth: "50px",
                        padding: "0",
                        marginRight: "1px",
                        marginLeft: "1px",
                    },

                    linkIsSelected: {
                        fontSize: "14px",
                        fontWeight: "bold",
                        color: "#107c41",
                        backgroundColor: "transparent",
                        width: "50px",
                        minWidth: "50px",
                        padding: "0",
                        selectors: {
                            // Đưa đường gạch xanh (khi được chọn) lên trên
                            "&::before": {
                                top: 0,                   // Ép dính vào mép trên cùng
                                bottom: "auto",           // Hủy bỏ vị trí ở dưới mặc định
                                height: "3px",            // (Tùy chọn) Chỉnh độ dày của đường kẻ cho đẹp
                                backgroundColor: "#107c41" // Đảm bảo giữ màu xanh chuẩn
                            }
                        }
                    }
                }}>

                {/* TAB 1: ĐẠI SỐ CƠ BẢN */}
                <PivotItem headerText="Cơ bản"
                    onRenderItemLink={() => renderLatexTab("\\frac{#?}{#?}")}>
                    <OverflowMathGroup>
                        <MathBtn displayLatex="\frac{#0}{#?}" insertLatex="\frac{#0}{#?}" title="\frac" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}/{#?}" insertLatex="{#0}/{#?}" title="/" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}\dfrac{#?}{#?}" insertLatex="{#0}\dfrac{#?}{#?}" title="{}\dfrac{}{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <VerticalDivider />
                        <MathBtn displayLatex="\sqrt{#0}" insertLatex="\sqrt{#0}" title="\sqrt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt[3]{#0}" insertLatex="\sqrt[3]{#0}" title="\sqrt[3]{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqrt[#?]{#0}" insertLatex="\sqrt[#?]{#0}" title="\sqrt[n]{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <VerticalDivider />
                        <MathBtn displayLatex="{#?}^{#?}" insertLatex="{#0}^{#?}" title="^{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}_{#?}" insertLatex="{#0}_{#?}" title="_{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="{#0}_{#?}^{#?}" insertLatex="{#0}_{#?}^{#?}" title="{}_{}^{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="_{#0}^{#?}{#?}" insertLatex="_{#0}^{#?}{#?}" title="_{}^{}{}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <VerticalDivider />
                        <MathBtn displayLatex="\left|#0\right|" insertLatex="\left|#0\right|" title="\left|\right|" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overline{#0}" insertLatex="\overline{#0}" title="\overline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underline{#0}" insertLatex="\underline{#0}" title="\underline" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </OverflowMathGroup>
                </PivotItem>
                {/* Operator */}
                <PivotItem headerText="Operator"
                    onRenderItemLink={() => renderLatexTab("\\pm")}>
                    <OverflowMathGroup>
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
                        <MathBtn displayLatex="\lt" insertLatex="\lt" title="\lt" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\le" insertLatex="\le" title="\le" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leqslant" insertLatex="\leqslant" title="\leqslant" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\propto" insertLatex="\propto" title="\propto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varpropto" insertLatex="\varpropto" title="\varpropto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangleq" insertLatex="\triangleq" title="\triangleq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\neq" insertLatex="\neq" title="\neq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\infty" insertLatex="\infty" title="\infty" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <VerticalDivider />

                        <MathBtn displayLatex="\ast" insertLatex="\ast" title="\ast" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\star" insertLatex="\star" title="\star" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\dagger" insertLatex="\dagger" title="\dagger" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ddagger" insertLatex="\ddagger" title="\ddagger" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cap" insertLatex="\cap" title="\cap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cup" insertLatex="\cup" title="\cup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uplus" insertLatex="\uplus" title="\uplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqcap" insertLatex="\sqcap" title="\sqcap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqcup" insertLatex="\sqcup" title="\sqcup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vee" insertLatex="\vee" title="\vee" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\wedge" insertLatex="\wedge" title="\wedge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdot" insertLatex="\cdot" title="\cdot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\diamond" insertLatex="\diamond" title="\diamond" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigtriangleup" insertLatex="\bigtriangleup" title="\bigtriangleup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangle" insertLatex="\triangle" title="\triangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigtriangledown" insertLatex="\bigtriangledown" title="\bigtriangledown" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangleleft" insertLatex="\triangleleft" title="\triangleleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\triangleright" insertLatex="\triangleright" title="\triangleright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigcirc" insertLatex="\bigcirc" title="\bigcirc" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bullet" insertLatex="\bullet" title="\bullet" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\wr" insertLatex="\wr" title="\wr" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oplus" insertLatex="\oplus" title="\oplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ominus" insertLatex="\ominus" title="\ominus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\otimes" insertLatex="\otimes" title="\otimes" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oslash" insertLatex="\oslash" title="\oslash" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\odot" insertLatex="\odot" title="\odot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\circ" insertLatex="\circ" title="\circ" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\setminus" insertLatex="\setminus" title="\setminus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\amalg" insertLatex="\amalg" title="\amalg" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </OverflowMathGroup>
                </PivotItem>

                {/* TAB 2: TOÁN TỬ & LOGIC */}
                <PivotItem headerText="Relation"
                    onRenderItemLink={() => renderLatexTab("\\ge")}>
                    <OverflowMathGroup>


                        <VerticalDivider />
                        <MathBtn displayLatex="<" insertLatex="<" title="<" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nless" insertLatex="\nless" title="\nless" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex=">" insertLatex=">" title=">" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ngtr" insertLatex="\ngtr" title="\ngtr" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\leq" insertLatex="\leq" title="\leq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nleq" insertLatex="\nleq" title="\nleq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\geq" insertLatex="\geq" title="\geq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ngeq" insertLatex="\ngeq" title="\ngeq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\leqslant" insertLatex="\leqslant" title="\leqslant" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nleqslant" insertLatex="\nleqslant" title="\nleqslant" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\geqslant" insertLatex="\geqslant" title="\geqslant" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ngeqslant" insertLatex="\ngeqslant" title="\ngeqslant" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\prec" insertLatex="\prec" title="\prec" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nprec" insertLatex="\nprec" title="\nprec" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\succ" insertLatex="\succ" title="\succ" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsucc" insertLatex="\nsucc" title="\nsucc" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\preceq" insertLatex="\preceq" title="\preceq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\npreceq" insertLatex="\npreceq" title="\npreceq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\succeq" insertLatex="\succeq" title="\succeq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsucceq" insertLatex="\nsucceq" title="\nsucceq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\ll" insertLatex="\ll" title="\ll" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\gg" insertLatex="\gg" title="\gg" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\lll" insertLatex="\lll" title="\lll" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ggg" insertLatex="\ggg" title="\ggg" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\subset" insertLatex="\subset" title="\subset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\not\subset" insertLatex="\not\subset" title="\not\subset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\supset" insertLatex="\supset" title="\supset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\not\supset" insertLatex="\not\supset" title="\not\supset" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\subseteq" insertLatex="\subseteq" title="\subseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsubseteq" insertLatex="\nsubseteq" title="\nsubseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\supseteq" insertLatex="\supseteq" title="\supseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsupseteq" insertLatex="\nsupseteq" title="\nsupseteq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\sqsubset" insertLatex="\sqsubset" title="\sqsubset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqsubseteq" insertLatex="\sqsubseteq" title="\sqsubseteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqsupset" insertLatex="\sqsupset" title="\sqsupset" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqsupseteq" insertLatex="\sqsupseteq" title="\sqsupseteq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="=" insertLatex="=" title="=" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\doteq" insertLatex="\doteq" title="\doteq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\equiv" insertLatex="\equiv" title="\equiv" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\approx" insertLatex="\approx" title="\approx" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cong" insertLatex="\cong" title="\cong" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\simeq" insertLatex="\simeq" title="\simeq" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sim" insertLatex="\sim" title="\sim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\propto" insertLatex="\propto" title="\propto" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\ne" insertLatex="\ne" title="\ne" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\neq" insertLatex="\neq" title="\neq" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\parallel" insertLatex="\parallel" title="\parallel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nparallel" insertLatex="\nparallel" title="\nparallel" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\vdash" insertLatex="\vdash" title="\vdash" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\dashv" insertLatex="\dashv" title="\dashv" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\in" insertLatex="\in" title="\in" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ni" insertLatex="\ni" title="\ni" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\notin" insertLatex="\notin" title="\notin" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\smile" insertLatex="\smile" title="\smile" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\frown" insertLatex="\frown" title="\frown" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\asymp" insertLatex="\asymp" title="\asymp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bowtie" insertLatex="\bowtie" title="\bowtie" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\models" insertLatex="\models" title="\models" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\perp" insertLatex="\perp" title="\perp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bot" insertLatex="\bot" title="\bot" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\mid" insertLatex="\mid" title="\mid" onInsert={onInsert} borderColor={PaletteColors.default} />

                    </OverflowMathGroup>
                </PivotItem>

                {/* TẬP HỢP */}
                <PivotItem headerText="Sets"
                    onRenderItemLink={() => renderLatexTab("\\forall")}>
                    <OverflowMathGroup>
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
                    </OverflowMathGroup>
                </PivotItem>

                {/* GIẢI TÍCH */}
                <PivotItem headerText="Giải tích"
                    onRenderItemLink={() => renderLatexTab("\\textstyle\\int")}>
                    <OverflowMathGroup>
                        <MathBtn displayLatex="\small\int_{#0}^{#1}" insertLatex="\int_{#0}^{#1}" title="\int" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="#?\Big|_#?^#?" insertLatex="#0\Big|_#?^#?" title="{}\Big_{}^{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\iint" insertLatex="\iint_{#0}^{#?}" title="\iint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\oint" insertLatex="\oint" title="\oint" onInsert={onInsert} borderColor={PaletteColors.default} />
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
                        <VerticalDivider />
                        <MathBtn displayLatex="\small\sum_{n=0}^{\infty}" insertLatex="\sum_{n=0}^{\infty}" title="\sum" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\prod_{n=0}^{\infty}" insertLatex="\prod_{n=0}^{\infty}" title="\prod" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\coprod_{n=0}^{\infty}" insertLatex="\coprod_{n=0}^{\infty}" title="\coprod" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\int_{0}^{\infty}" insertLatex="\int_{0}^{\infty}" title="\int" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\intop_{0}^{\infty}" insertLatex="\intop_{0}^{\infty}" title="\intop" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\iint_{0}^{\infty}" insertLatex="\iint_{0}^{\infty}" title="\iint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\iiint_{0}^{\infty}" insertLatex="\iiint_{0}^{\infty}" title="\iiint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\small\oint_{C}" insertLatex="\oint_{C}" title="\oint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\smallint" insertLatex="\smallint" title="\smallint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigcup" insertLatex="\bigcup" title="\bigcup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigcap" insertLatex="\bigcap" title="\bigcap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigvee" insertLatex="\bigvee" title="\bigvee" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigwedge" insertLatex="\bigwedge" title="\bigwedge" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\biguplus" insertLatex="\biguplus" title="\biguplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigotimes" insertLatex="\bigotimes" title="\bigotimes" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigoplus" insertLatex="\bigoplus" title="\bigoplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigodot" insertLatex="\bigodot" title="\bigodot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bigsqcup" insertLatex="\bigsqcup" title="\bigsqcup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oiint" insertLatex="\oiint" title="\oiint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\oiiint" insertLatex="\oiiint" title="\oiiint" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\intclockwise" insertLatex="\intclockwise" title="\intclockwise" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\varointclockwise" insertLatex="\varointclockwise" title="\varointclockwise" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ointctrclockwise" insertLatex="\ointctrclockwise" title="\ointctrclockwise" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\intctrclockwise" insertLatex="\intctrclockwise" title="\intctrclockwise" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Cap" insertLatex="\Cap" title="\Cap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Cup" insertLatex="\Cup" title="\Cup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\doublecap" insertLatex="\doublecap" title="\doublecap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\doublecup" insertLatex="\doublecup" title="\doublecup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqcup" insertLatex="\sqcup" title="\sqcup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sqcap" insertLatex="\sqcap" title="\sqcap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uplus" insertLatex="\uplus" title="\uplus" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\wr" insertLatex="\wr" title="\wr" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\amalg" insertLatex="\amalg" title="\amalg" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </OverflowMathGroup>
                </PivotItem>

                {/* TAB 5: HỆ & MA TRẬN - ĐÃ ĐƯỢC CẬP NHẬT */}
                <PivotItem headerText="Hệ & Ma trận"
                    onRenderItemLink={() => renderLatexTab("\\begin{bmatrix} a &b \\\\ c &d \\end{bmatrix}")}>
                    <OverflowMathGroup>
                        {/* CÁC NÚT HỆ PHƯƠNG TRÌNH / HÀM SỐ ĐÃ ĐƯỢC CHUYỂN SANG ĐỘNG */}
                        <SystemSelectorBtn displayLatex="\lbrace\text{■}" type="cases" title="\begin{cases}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <SystemSelectorBtn displayLatex="\left[ \text{■}" type="array" title="\left[ \begin{array}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <SystemSelectorBtn displayLatex="\left[ \text{m}" type="matrix" title="\left[ \begin{matrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="f(x) = \lbrace" insertLatex="f(x) = \begin{cases} #?, & \text{nếu} #? \\ #?, & \text{nếu} #? \end{cases}" title="f(x) \begin{cases}" onInsert={onInsert} borderColor={PaletteColors.default} />

                        {/* CÁC NÚT MA TRẬN ĐỘNG */}
                        <MatrixSelectorBtn displayLatex="\text{(■)}" matrixType="pmatrix" title="\begin{pmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MatrixSelectorBtn displayLatex="\text{[■]}" matrixType="bmatrix" title="\begin{bmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MatrixSelectorBtn displayLatex="\text{|■|}" matrixType="vmatrix" title="\begin{vmatrix}" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <VerticalDivider />
                        <MathBtn displayLatex="\binom{#?}{#?}" insertLatex="\binom{#0}{#?}" title="\binom{}{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\dbinom{#?}{#?}" insertLatex="\dbinom{#0}{#?}" title="\dbinom{}{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\tbinom{#?}{#?}" insertLatex="\tbinom{#0}{#?}" title="\tbinom{}{}" onInsert={onInsert} borderColor={PaletteColors.default} />

                        {/* Các nút phụ tĩnh trợ năng giữ nguyên */}
                        <MathBtn displayLatex="\text{[⋮]}" insertLatex="\begin{bmatrix} #? \\ #? \\ #? \end{bmatrix}" title="\begin{bmatrix} 3x1" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cdots" insertLatex="\cdots" title="\cdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vdots" insertLatex="\vdots" title="\vdots" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ddots" insertLatex="\ddots" title="\ddots" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </OverflowMathGroup>
                </PivotItem>

                {/* HÌNH & LƯỢNG GIÁC */}
                <PivotItem headerText="Lượng giác"
                    onRenderItemLink={() => renderLatexTab("\\sin(\\pi)")}>
                    <OverflowMathGroup>
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
                        <MathBtn displayLatex="\degree" insertLatex="\degree" title="\degree" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\angle" insertLatex="\angle" title="\angle" onInsert={onInsert} borderColor={PaletteColors.default} />
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
                        <VerticalDivider />
                        <MathBtn displayLatex="\overline{AB}" insertLatex="\overline{AB}" title="\overline{}" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\bigtriangleup" insertLatex="\bigtriangleup" title="\bigtriangleup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cong" insertLatex="\cong" title="\cong" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\ncong" insertLatex="\ncong" title="\ncong" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\sim" insertLatex="\sim" title="\sim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nsim" insertLatex="\nsim" title="\nsim" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\|" insertLatex="\|" title="\|" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\nparallel" insertLatex="\nparallel" title="\nparallel" onInsert={onInsert} borderColor={PaletteColors.default} />

                        <MathBtn displayLatex="\bot" insertLatex="\bot" title="\bot" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\not\perp" insertLatex="\not\perp" title="\not\perp" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overrightarrow{AB}" insertLatex="\overrightarrow{AB}" title="\overrightarrow{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overleftarrow{ABC}" insertLatex="\overleftarrow{\ABC}" title="\overleftarrow{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\vec{a}" insertLatex="\vec{a}" title="\vec{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\measuredangle" insertLatex="\measuredangle" title="\measuredangle" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\square" insertLatex="\square" title="\square" onInsert={onInsert} borderColor={PaletteColors.default} />
                    </OverflowMathGroup>
                </PivotItem>

                {/* LÝ & HÓA HỌC */}
                <PivotItem headerText="Lý & Hóa"
                    onRenderItemLink={() => renderLatexTab("t^{\\circ} \\, ^{A}_{Z}X")}>
                    <OverflowMathGroup>
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow" title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightleftharpoons" insertLatex="\rightleftharpoons" title="\rightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uparrow" insertLatex="\uparrow" title="\uparrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downarrow" insertLatex="\downarrow" title="\downarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="t^\circ" insertLatex="t^\circ" title="t^\circ" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overset{t^\circ}{\rightarrow}" insertLatex="\overset{t^\circ}{\rightarrow}" title="\overset{t^\circ}{\rightarrow}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xrightarrow{\text{xt}}" insertLatex="\xrightarrow{\text{xt}}" title="\xrightarrow{\text{xt}}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="e^{-}" insertLatex="e^{-}" title="e^{-}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="^{A}_{Z}X" insertLatex="^{#?}_{#?}#?}" title="^{A}_{Z}X}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\Omega" insertLatex="\Omega" title="\Omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\omega" insertLatex="\omega" title="\omega" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\lambda" insertLatex="\lambda" title="\lambda" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mu" insertLatex="\mu" title="\mu" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\text{\AA}" insertLatex="\text{\AA}" title="\text{\AA}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hbar" insertLatex="\hbar" title="\hbar" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\mathscr{E}" insertLatex="\mathscr{E}" title="\mathscr{E}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <VerticalDivider />
                        <MathBtn displayLatex="\ce{#? -> #?}" insertLatex="\ce{#0 -> #?}" title="\ce{#? -> #?}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\ce{#? ->[a] #?}" insertLatex="\ce{#0 ->[#?] #?}" title="\ce{#? ->[#?] #?}" onInsert={onInsert} borderColor={PaletteColors.latex} />

                    </OverflowMathGroup>
                </PivotItem>

                {/* Accents*/}
                <PivotItem headerText="Accents"
                    onRenderItemLink={() => renderLatexTab("\\dddot{#?}")}>
                    <OverflowMathGroup>
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
                        <MathBtn displayLatex="\dot{#?}" insertLatex="\dot{#0}" title="\dot" onInsert={onInsert} borderColor={PaletteColors.office} />

                        <MathBtn displayLatex="\ddot{#?}" insertLatex="\ddot{#0}" title="\ddot" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\dddot{#?}" insertLatex="\dddot{#0}" title="\dddot" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\ddddot{#?}" insertLatex="\ddddot{#0}" title="\ddddot" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\mathring{#?}" insertLatex="\mathring{#0}" title="\mathring" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\tilde{#?}" insertLatex="\tilde{#0}" title="\tilde" onInsert={onInsert} borderColor={PaletteColors.office} />

                        <MathBtn displayLatex="\breve{#?}" insertLatex="\breve{#0}" title="\breve" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\check{#?}" insertLatex="\check{#0}" title="\check" onInsert={onInsert} borderColor={PaletteColors.office} />

                    </OverflowMathGroup>
                </PivotItem>

                {/* MŨI TÊN */}
                <PivotItem headerText="Arrows"
                    onRenderItemLink={() => renderLatexTab("\\rightarrow")}>
                    <OverflowMathGroup>
                        <MathBtn displayLatex="\leftarrow" insertLatex="\leftarrow" title="\leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightarrow" insertLatex="\rightarrow" title="\rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightarrow" insertLatex="\leftrightarrow" title="\leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftarrow" insertLatex="\Leftarrow" title="\Leftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Rightarrow" insertLatex="\Rightarrow" title="\Rightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Leftrightarrow" insertLatex="\Leftrightarrow" title="\Leftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downarrow" insertLatex="\downarrow" title="\downarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\uparrow" insertLatex="\uparrow" title="\uparrow" onInsert={onInsert} borderColor={PaletteColors.default} />
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
                        <VerticalDivider />
                        
                        
                        <MathBtn displayLatex="\twoheadrightarrow" insertLatex="\twoheadrightarrow" title="\twoheadrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\twoheadleftarrow" insertLatex="\twoheadleftarrow" title="\twoheadleftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightarrowtail" insertLatex="\rightarrowtail" title="\rightarrowtail" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftarrowtail" insertLatex="\leftarrowtail" title="\leftarrowtail" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\dashrightarrow" insertLatex="\dashrightarrow" title="\dashrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\dashleftarrow" insertLatex="\dashleftarrow" title="\dashleftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\longrightarrow" insertLatex="\longrightarrow" title="\longrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\longleftarrow" insertLatex="\longleftarrow" title="\longleftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\longleftrightarrow" insertLatex="\longleftrightarrow" title="\longleftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        
                        <MathBtn displayLatex="\Longrightarrow" insertLatex="\Longrightarrow" title="\Longrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\Longleftarrow" insertLatex="\Longleftarrow" title="\Longleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\Longleftrightarrow" insertLatex="\Longleftrightarrow" title="\Longleftrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\longmapsto" insertLatex="\longmapsto" title="\longmapsto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\multimap" insertLatex="\multimap" title="\multimap" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        
                        <MathBtn displayLatex="\updownarrow" insertLatex="\updownarrow" title="\updownarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Uparrow" insertLatex="\Uparrow" title="\Uparrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Downarrow" insertLatex="\Downarrow" title="\Downarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Updownarrow" insertLatex="\Updownarrow" title="\Updownarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightharpoonup" insertLatex="\rightharpoonup" title="\rightharpoonup" onInsert={onInsert} borderColor={PaletteColors.default} />
                        
                        <MathBtn displayLatex="\rightharpoondown" insertLatex="\rightharpoondown" title="\rightharpoondown" onInsert={onInsert} borderColor={PaletteColors.default} />
                       
                        <MathBtn displayLatex="\rightleftharpoons" insertLatex="\rightleftharpoons" title="\rightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightharpoons" insertLatex="\leftrightharpoons" title="\leftrightharpoons" onInsert={onInsert} borderColor={PaletteColors.default} />                       
                        <MathBtn displayLatex="\Rrightarrow" insertLatex="\Rrightarrow" title="\Rrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Lleftarrow" insertLatex="\Lleftarrow" title="\Lleftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightarrows" insertLatex="\leftrightarrows" title="\leftrightarrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightleftarrows" insertLatex="\rightleftarrows" title="\rightleftarrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\curvearrowright" insertLatex="\curvearrowright" title="\curvearrowright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\curvearrowleft" insertLatex="\curvearrowleft" title="\curvearrowleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hookrightarrow" insertLatex="\hookrightarrow" title="\hookrightarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\hookleftarrow" insertLatex="\hookleftarrow" title="\hookleftarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\looparrowright" insertLatex="\looparrowright" title="\looparrowright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\looparrowleft" insertLatex="\looparrowleft" title="\looparrowleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\circlearrowright" insertLatex="\circlearrowright" title="\circlearrowright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\circlearrowleft" insertLatex="\circlearrowleft" title="\circlearrowleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightrightarrows" insertLatex="\rightrightarrows" title="\rightrightarrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftleftarrows" insertLatex="\leftleftarrows" title="\leftleftarrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\upuparrows" insertLatex="\upuparrows" title="\upuparrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downdownarrows" insertLatex="\downdownarrows" title="\downdownarrows" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Rsh" insertLatex="\Rsh" title="\Rsh" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\Lsh" insertLatex="\Lsh" title="\Lsh" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\upharpoonright" insertLatex="\upharpoonright" title="\upharpoonright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\upharpoonleft" insertLatex="\upharpoonleft" title="\upharpoonleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downharpoonright" insertLatex="\downharpoonright" title="\downharpoonright" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\downharpoonleft" insertLatex="\downharpoonleft" title="\downharpoonleft" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\restriction" insertLatex="\restriction" title="\restriction" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\rightsquigarrow" insertLatex="\rightsquigarrow" title="\rightsquigarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leftrightsquigarrow" insertLatex="\leftrightsquigarrow" title="\leftrightsquigarrow" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\leadsto" insertLatex="\leadsto" title="\leadsto" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <VerticalDivider />
                        {/* Hàng 1 */}
                       
                    </OverflowMathGroup>
                </PivotItem>

                {/* Symbols */}
                <PivotItem headerText="Symbols"
                    onRenderItemLink={() => renderLatexTab("\\infty")}>
                    <OverflowMathGroup>
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
                        <MathBtn displayLatex="\Im" insertLatex="\Im" title="\Im" onInsert={onInsert} borderColor={PaletteColors.office} />
                        <MathBtn displayLatex="\Re" insertLatex="\Re" title="\Re" onInsert={onInsert} borderColor={PaletteColors.office} />
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
                    </OverflowMathGroup>
                </PivotItem>

                {/* Annotation */}
                <PivotItem headerText="Annotation"
                    onRenderItemLink={() => renderLatexTab("\\overbrace{#?}^{\\text{note}}")}>
                    <OverflowMathGroup>
                        <MathBtn displayLatex="\text{txt}" insertLatex="\text{#0}" title="\text{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\textrm{txt}" insertLatex="\textrm{#0}" title="\textrm{}" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\cancel{2}" insertLatex="\cancel{#0}" title="\cancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\bcancel{2}" insertLatex="\bcancel{#0}" title="\bcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\xcancel{ABC}" insertLatex="\xcancel{ABC}" title="\xcancel" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\overbrace{a+b+c}^{\text{note}}" insertLatex="\overbrace{a+b+c}^{\text{#0}}" title="\overbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\underbrace{a+b+c}_{\text{note}}" insertLatex="\underbrace{a+b+c}_{\text{#0}}" title="\underbrace" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\boxed{AB}" insertLatex="\boxed{#0}" title="\boxed" onInsert={onInsert} borderColor={PaletteColors.default} />
                        <MathBtn displayLatex="\phase{60^\circ}" insertLatex="\phase{#0}" title="\phase{}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\enclose{roundedbox updiagonalstrike}{x=0}" insertLatex="\enclose{roundedbox updiagonalstrike}{#0}" title="\enclose{roundedbox updiagonalstrike}{}" onInsert={onInsert} borderColor={PaletteColors.office} />
                         <MathBtn displayLatex="\angl{#?}" insertLatex="\angl{#0}" title="\angl{}" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <VerticalDivider />
                        <MathBtn displayLatex="\xrightarrow[\text{long text below}]{}" insertLatex="\xrightarrow[#?]{#0}" title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightarrow{\text{long text above}}" insertLatex="\xrightarrow{#0}" title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightarrow[\text{and below}]{\text{long text above}}" insertLatex="\xrightarrow[#?]{#0}" title="\xrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xlongequal[a]{b}" insertLatex="\xlongequal[#?]{#0}" title="\xlongequal" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftarrow[a]{b}" insertLatex="\xleftarrow[#?]{#0}" title="\xleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightarrow[a]{b}" insertLatex="\xleftrightarrow[#?]{#0}" title="\xleftrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xtwoheadrightarrow[a]{b}" insertLatex="\xtwoheadrightarrow[#?]{#0}" title="\xtwoheadrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xtwoheadleftarrow[a]{b}" insertLatex="\xtwoheadleftarrow[#?]{#0}" title="\xtwoheadleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xRightarrow[a]{b}" insertLatex="\xRightarrow[#?]{#0}" title="\xRightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xLeftarrow[a]{b}" insertLatex="\xLeftarrow[#?]{#0}" title="\xLeftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightharpoonup[a]{b}" insertLatex="\xrightharpoonup[#?]{#0}" title="\xrightharpoonup" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftharpoonup[a]{b}" insertLatex="\xleftharpoonup[#?]{#0}" title="\xleftharpoonup" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightharpoondown[a]{b}" insertLatex="\xrightharpoondown[#?]{#0}" title="\xrightharpoondown" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftharpoondown[a]{b}" insertLatex="\xleftharpoondown[#?]{#0}" title="\xleftharpoondown" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xrightleftharpoons[a]{b}" insertLatex="\xrightleftharpoons[#?]{#0}" title="\xrightleftharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xleftrightharpoons[a]{b}" insertLatex="\xleftrightharpoons[#?]{#0}" title="\xleftrightharpoons" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xhookrightarrow[a]{b}" insertLatex="\xhookrightarrow[#?]{#0}" title="\xhookrightarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xhookleftarrow[a]{b}" insertLatex="\xhookleftarrow[#?]{#0}" title="\xhookleftarrow" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xmapsto[a]{b}" insertLatex="\xmapsto[#?]{#0}" title="\xmapsto" onInsert={onInsert} borderColor={PaletteColors.latex} />
                        <MathBtn displayLatex="\xtofrom[a]{b}" insertLatex="\xtofrom[#?]{#0}" title="\xtofrom" onInsert={onInsert} borderColor={PaletteColors.latex} />
                    </OverflowMathGroup>
                </PivotItem>

                {/* TAB 9: HY LẠP */}
                <PivotItem headerText="Greek letters"
                    onRenderItemLink={() => renderLatexTab("\\alpha")}>
                    <OverflowMathGroup>
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
                    </OverflowMathGroup>
                </PivotItem>

            </Pivot>
        </div>
    );
}