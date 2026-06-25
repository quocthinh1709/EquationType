// src/taskpane/components/mathShortcuts.ts

// 1. Danh sách phím tắt gõ nhanh (Inline Shortcuts)
export const customInlineShortcuts = {
    // Phân số, căn, vector
    "ps": "\\frac{#0}{#?}",
    "vt": "\\vec{#0}",
    "can": "\\sqrt{#0}",
    
    // Logic, toán tử
    "td": "\\Leftrightarrow ",
    "sr": "\\Rightarrow ",
    "thuoc": "\\in ",
    "<=>": "\\Leftrightarrow ",
    "->": "\\rightarrow ",

    // Hệ phương trình
    "hept2an": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\end{cases}",
    "hept3an": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\\\ #? &= #? \\end{cases}",
    
    // Tích phân, đạo hàm
    "tp": "\\int_{#?}^{#?}",
    "dh": "\\frac{d}{dx} #0",

    "tp3": "∰",        
    "song": "⤳", 
    "gvuong": "⦜",
    "angstrom": "Å",
    "male": "♂",
    "female": "♀",
};

// 2. Danh sách lệnh tùy chỉnh (Macros - gõ có dấu \ phía trước)
export const customMacros = {
    // Các tập số học
    "tapR": "\\mathbb{R}",
    "tapN": "\\mathbb{N}",
    "tapZ": "\\mathbb{Z}",
    "tapQ": "\\mathbb{Q}",
    // Phân số, căn, vector
    "phanso": "\\frac{#0}{#?}",
    "vecto": "\\vec{#0}",
    "canbac2": "\\sqrt{#0}",
    
    // Logic, toán tử
    "td": "\\Leftrightarrow ",
    "suyra": "\\Rightarrow ",
    "thuoc": "\\in ",
    "tuongduong": "\\Leftrightarrow ",
    "->": "\\rightarrow ",

    // Hệ phương trình
    "hehaipt": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\end{cases}",
    "hebapt": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\\\ #? &= #? \\end{cases}",
    
    // Tích phân, đạo hàm
    "ticphan": "\\int_{#?}^{#?}",
    "daoham": "\\frac{d}{dx} #0",
    "male": "♂",
    "female": "♀",
    "angstrom": "Å",
  
};