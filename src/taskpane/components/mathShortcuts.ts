// src/taskpane/components/mathShortcuts.ts

// 1. Danh sách phím tắt gõ nhanh (Inline Shortcuts)
export const customInlineShortcuts = {
    // Phân số, căn, vector
    "ps": "\\frac{#0}{#?}",
    "vt": "\\vec{#0}",
    "can": "\\sqrt{#0}",
    "|": "\\mid",
    
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

    "tp3": "\\iiint",        
    "song": "\\rightsquigarrow", 
    "gvuong": "\\square",
    "angstrom": "\\text{Å}",
    "male": "\\text{♂}",
    "female": "\\text{♀}",

    // ==========================================
    // TOÁN HỌC - CÁC BIỂU THỨC LỚN
    // ==========================================
    // Giải tích
    "dh11": "y' = \\frac{#? \\cdot #? - #? \\cdot #?}{(#? x + #?)^2}", // Đạo hàm bậc 1/bậc 1: y' = (ad-bc)/(cx+d)^2
    "tptp": "\\int #? \\, d(#?) = #? \\cdot #? - \\int #? \\, d(#?)",  // Tích phân từng phần (udv = uv - vdu)
    "dtphang": "S = \\int_{#?}^{#?} \\left| #? - #? \\right| dx",      // Diện tích hình phẳng

    // Hình học & Đại số
    "ptb2": "x_{1,2} = \\frac{-#? \\pm \\sqrt{#?}}{2 \\cdot #?}",      // Công thức nghiệm phương trình bậc 2
    "khoangcach": "d(M, (P)) = \\frac{\\left| #? \\right|}{\\sqrt{#?^2 + #?^2 + #?^2}}", // Khoảng cách từ điểm đến mặt phẳng
    "ptduongthang": "\\frac{x - #?}{#?} = \\frac{y - #?}{#?} = \\frac{z - #?}{#?}",      // Phương trình đường thẳng (chính tắc)
    
    // Lượng giác & Nhị thức Newton
    "sin2x": "\\sin 2x = 2\\sin x \\cos x",
    "cos2x": "\\cos 2x = \\cos^2 x - \\sin^2 x",
    "nhithuc": "(a+b)^n = \\sum_{k=0}^{n} C_n^k a^{n-k} b^k",
    "vchop": "V = \\frac{1}{3} B \\cdot h",              // Thể tích khối chóp
    "vltru": "V = B \\cdot h",                           // Thể tích khối lăng trụ
    "vcau": "V = \\frac{4}{3}\\pi R^3",                  // Thể tích khối cầu
    "scau": "S = 4\\pi R^2",                             // Diện tích mặt cầu
    "vtron": "V = \\pi R^2 h",                           // Thể tích khối trụ tròn xoay
    "vnon": "V = \\frac{1}{3}\\pi R^2 h",                // Thể tích khối nón
    "ptchan": "\\frac{x}{a} + \\frac{y}{b} + \\frac{z}{c} = 1", // Phương trình mặt phẳng theo đoạn chắn
    "modun": "\\left| z \\right| = \\sqrt{a^2 + b^2}",    // Môđun số phức z = a + bi

    // Dãy số & Tiến trình
    "csc": "u_n = u_1 + (n-1)d",                         // Cấp số cộng
    "sncsc": "S_n = \\frac{n(u_1 + u_n)}{2}",            // Tổng n số hạng cấp số cộng
    "csn": "u_n = u_1 \\cdot q^{n-1}",                   // Cấp số nhân
    // Logic & Tập hợp
    "vmo": "\\forall #? \\in #?",                     // Với mọi
    "ttai": "\\exists #? \\in #?",                    // Tồn tại
    "rong": "\\varnothing",                           // Tập rỗng
    "giao": "\\cap",                                  // Giao của 2 tập hợp
    "hop": "\\cup",                                   // Hợp của 2 tập hợp
    "hieu": "\\setminus",                             // Hiệu của 2 tập hợp
    
    // Hằng đẳng thức & Giá trị
    "hdt1": "(#? + #?)^2 = #?^2 + 2\\cdot#?\\cdot#? + #?^2", // Hằng đẳng thức 1
    "hdt2": "(#? - #?)^2 = #?^2 - 2\\cdot#?\\cdot#? + #?^2", // Hằng đẳng thức 2
    "hdt3": "#?^2 - #?^2 = (#? - #?)(#? + #?)",              // Hằng đẳng thức 3
    "gttd": "\\left| #? \\right|",                           // Giá trị tuyệt đối
    "cbn": "\\sqrt[#?]{#?}",                                 // Căn bậc n

    // Hệ thức lượng & Lượng giác
    "dlcos": "#?^2 = #?^2 + #?^2 - 2\\cdot#?\\cdot#?\\cos #?", // Định lý Cosin
    "dlsin": "\\frac{#?}{\\sin #?} = 2R",                      // Định lý Sin
    "ptlg1": "x = \\alpha + k2\\pi \\lor x = \\pi - \\alpha + k2\\pi", // Nghiệm sinx = sin alpha
    
    // ==========================================
    // TOÁN HỌC (Math) - GIẢI TÍCH & HÌNH HỌC OXYZ
    // ==========================================
    "dh2": "y'' = #?",                                       // Đạo hàm bậc 2
    "txd": "D = \\mathbb{R} \\setminus \\{ #? \\}",          // Tập xác định
    "tch": "\\left[ \\vec{#?}, \\vec{#?} \\right]",          // Tích có hướng 2 vector
    "ptmp": "#?(x - #?) + #?(y - #?) + #?(z - #?) = 0",      // Phương trình mặt phẳng
    "ptmc": "(x - #?)^2 + (y - #?)^2 + (z - #?)^2 = R^2",    // Phương trình mặt cầu
    "toado": "#?(#?; #?; #?)",                               // Tọa độ điểm M(x;y;z)

    // ==========================================
    // VẬT LÝ - CÁC PHƯƠNG TRÌNH ĐẶC TRƯNG
    // ==========================================
    "ptcd": "x = x_0 + v_0 t + \\frac{1}{2} a t^2",                     // Phương trình chuyển động thẳng biến đổi đều
    "cklx": "T = 2\\pi \\sqrt{\\frac{m}{k}}",                           // Chu kỳ con lắc lò xo
    "ckdd": "T = 2\\pi \\sqrt{\\frac{l}{g}}",                           // Chu kỳ con lắc đơn
    "dlom": "I = \\frac{\\mathcal{E}}{R_N + r}",                        // Định luật Ohm cho toàn mạch
    "ptdao": "x = A \\cos(\\omega t + \\varphi)",                       // Phương trình dao động điều hòa
    "htquang": "\\frac{1}{d} + \\frac{1}{d'} = \\frac{1}{f}",           // Công thức thấu kính

    "bsong": "\\lambda = v \\cdot T = \\frac{v}{f}",     // Bước sóng
    "vsang": "x = k \\frac{\\lambda D}{a}",              // Vị trí vân sáng (Giao thoa ánh sáng)
    "vtoi": "x = \\left(k + \\frac{1}{2}\\right) \\frac{\\lambda D}{a}", // Vị trí vân tối
    "itruong": "i = \\frac{\\lambda D}{a}",              // Khoảng vân i
    "tthong": "\\Phi = B \\cdot S \\cdot \\cos \\alpha", // Từ thông
    "sddcu": "e_c = -\\frac{\\Delta \\Phi}{\\Delta t}", // Suất điện động cảm ứng
    "fcl": "F = q \\cdot E",                             // Lực điện trường
    "flen": "F = |q| v B \\sin \\alpha",                 // Lực Lorentz

        // Cơ học
    "dl2nw": "\\vec{F} = m\\vec{a}",                         // Định luật 2 Newton
    "fms": "F_{ms} = \\mu N",                                // Lực ma sát
    "fdh": "F_{dh} = -k \\Delta l",                          // Lực đàn hồi
    "fhd": "F_{hd} = G \\frac{m_1 m_2}{r^2}",                // Lực hấp dẫn
    "btcd": "p = mv",                                        // Động lượng

    // Điện xoay chiều
    "tongtro": "Z = \\sqrt{R^2 + (Z_L - Z_C)^2}",            // Tổng trở Z
    "dolechpha": "\\tan \\varphi = \\frac{Z_L - Z_C}{R}",    // Độ lệch pha
    "congsuat": "P = UI \\cos \\varphi = I^2 R",             // Công suất tỏa nhiệt
    "hsdc": "k = \\cos \\varphi = \\frac{R}{Z}",             // Hệ số công suất

    // Lượng tử & Hạt nhân
    "nlpt": "\\varepsilon = hf = \\frac{hc}{\\lambda}",      // Năng lượng photon
    "ast": "hf = A + \\frac{1}{2}mv_{0\\max}^2",             // Hệ thức Einstein (quang điện)
    "dhk": "\\Delta m = [Z m_p + (A-Z) m_n] - m_X",          // Độ hụt khối
    "nllk": "W_{lk} = \\Delta m \\cdot c^2",                 // Năng lượng liên kết
    "dlpx": "N = N_0 \\cdot 2^{-\\frac{t}{T}} = N_0 e^{-\\lambda t}", // Định luật phóng xạ

    // ==========================================
    // HÓA HỌC - CÁC CÔNG THỨC TÍNH TOÁN
    // ==========================================
    "kltb": "\\overline{M} = \\frac{#? \\cdot #? + #? \\cdot #?}{100}", // Nguyên tử khối trung bình
    "khi": "pV = nRT",                                                  // Phương trình trạng thái khí lý tưởng
    "ph": "pH = -\\log \\left[ H^+ \\right]",                           // Công thức tính pH
    "bte": "\\sum n_{e(\\text{nhường})} = \\sum n_{e(\\text{nhận})}",   // Bảo toàn electron
    "btkl": "m_{\\text{trước}} = m_{\\text{sau}}",                      // Bảo toàn khối lượng
    "cpt": "C\\% = \\frac{m_{ct}}{m_{dd}} \\cdot 100\\%", // Nồng độ phần trăm
    "cmol": "C_M = \\frac{n}{V}",                        // Nồng độ mol
    "dkl": "D = \\frac{m}{V}",                           // Khối lượng riêng D = m/V
    "alkan": "\\text{C}_n\\text{H}_{2n+2}",             // Công thức tổng quát Ankan
    "alken": "\\text{C}_n\\text{H}_{2n}",               // Anken
    "alkin": "\\text{C}_n\\text{H}_{2n-2}",              // Ankin
    "ruou": "\\text{C}_n\\text{H}_{2n+1}\\text{OH}",     // Ancol đơn chức, no
    "axit": "\\text{C}_n\\text{H}_{2n+1}\\text{COOH}",   // Axit cacboxylic no, đơn chức
    "kcb": "K_c = \\frac{[#?]^#? \\cdot [#?]^#?}{[#?]^#? \\cdot [#?]^#?}", // Hằng số cân bằng Kc
    "faraday": "m = \\frac{A}{n} \\cdot \\frac{It}{F}",      // Định luật Faraday (Điện phân)
    "tst": "T = [#?^+]^#? \\cdot [#?^-]^#?",                 // Tích số tan
    "xni": "\\xrightarrow{Ni, t^\\circ}",                    // Xúc tác Niken, nhiệt độ
    "xh2so4": "\\xrightarrow{H_2SO_4 \\text{ đ}, t^\\circ}", // Xúc tác H2SO4 đặc
    "vongbz": "\\text{C}_6\\text{H}_5\\text{-}",             // Gốc phenyl (thay vì vẽ vòng phức tạp)

    // ==========================================
    // SINH HỌC - CÁC CÔNG THỨC DI TRUYỀN
    // ==========================================
    "hardy": "p^2 AA + 2pq Aa + q^2 aa = 1",                            // Quần thể Hardy-Weinberg
    "chieuadn": "L = \\frac{N}{2} \\times 3,4 \\text{ \\AA}",           // Chiều dài phân tử ADN
    "nuadn": "A + G = T + X = \\frac{N}{2}",                            // Nguyên tắc bổ sung ADN
    "lkhydro": "H = 2A + 3G",                            // Số liên kết Hydro trong ADN
    "ckxoan": "C = \\frac{N}{20}",                       // Số chu kì xoắn của gen
    "noo": "2^n",                                        // Số tế bào con sau n lần nguyên phân
    "aa": "\\frac{N}{6} - 1",                            // Số axit amin trong chuỗi polipeptit hoàn chỉnh
    "fhvg": "f = \\frac{\\text{Số cá thể hoán vị}}{\\text{Tổng số cá thể}} \\times 100\\%", // Tần số hoán vị gen
    "dotbien": "2n \\pm 1",                                  // Đột biến lệch bội
    "kthuocqt": "N_t = N_0 + B - D + I - E",                 // Kích thước quần thể sinh vật
    

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
    "canbachai": "\\sqrt{#0}",
    "canbacba": "\\sqrt[3]{#0}",
   
    // Logic, toán tử
    "td": "\\Leftrightarrow ",
    "suyra": "\\Rightarrow ",
    "thuoc": "\\in ",
    "tuongduong": "\\Leftrightarrow ",
    "->": "\\rightarrow ",

    // Hệ phương trình
    "hehaipt": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\end{cases}",
    "hebapt": "\\begin{cases} #0 &= #? \\\\ #? &= #? \\\\ #? &= #? \\end{cases}",
    "hehaiptnew": "\\left\\{ \\begin{aligned} #0 &= #? \\\\ #? &= #? \\end{aligned} \\right",
    
    // Tích phân, đạo hàm
    "ticphan": "\\int_{#?}^{#?}",
    "daoham": "\\frac{d}{dx} #0",
    "male": "\\text{♂}",
    "female": "\\text{♀}",
    "angstrom": "\\text{Å}",

    //
    "tamgiac": "\\triangle",
    "dongdang": "\\backsim",
    "dongdangvn": "\\text{ᔕ}",

    // ==========================================
    // TOÁN HỌC
    // ==========================================
    "daohambacnhat": "y' = \\frac{#? \\cdot #? - #? \\cdot #?}{(#? x + #?)^2}",
    "tichphantungphan": "\\int #? \\, d(#?) = #? \\cdot #? - \\int #? \\, d(#?)",
    "nghiembac2": "x_{1,2} = \\frac{-#? \\pm \\sqrt{#?}}{2 \\cdot #?}",
    "khoangcachmatphang": "d(M, (P)) = \\frac{\\left| #? \\right|}{\\sqrt{#?^2 + #?^2 + #?^2}}",
    "voimoi": "\\forall",
    "tontai": "\\exists",
    "taprong": "\\varnothing",
    "giaotaphop": "\\cap",
    "hoptaphop": "\\cup",
    "giatrituyetdoi": "\\left| #? \\right|",
    "canbacn": "\\sqrt[#?]{#?}",
    "dinhlycosin": "#?^2 = #?^2 + #?^2 - 2\\cdot#?\\cdot#?\\cos #?",
    "dinhlysin": "\\frac{#?}{\\sin #?} = 2R",
    "tichcohuong": "\\left[ \\vec{#?}, \\vec{#?} \\right]",
    "ptmatphang": "#?(x - #?) + #?(y - #?) + #?(z - #?) = 0",
    "ptmatcau": "(x - #?)^2 + (y - #?)^2 + (z - #?)^2 = R^2",
    "thetichchop": "V = \\frac{1}{3} B \\cdot h",
    "thetichlangtru": "V = B \\cdot h",
    "thetichcau": "V = \\frac{4}{3}\\pi R^3",
    "dientichcau": "S = 4\\pi R^2",
    "thetichtru": "V = \\pi R^2 h",
    "thetichnon": "V = \\frac{1}{3}\\pi R^2 h",
    "matphangdoanchan": "\\frac{x}{a} + \\frac{y}{b} + \\frac{z}{c} = 1",
    "modunsophuc": "\\left| z \\right| = \\sqrt{a^2 + b^2}",
    "capsocong": "u_n = u_1 + (n-1)d",
    "capsonhan": "u_n = u_1 \\cdot q^{n-1}",
    
    // ==========================================
    // VẬT LÝ
    // ==========================================
    "chuyendong": "x = x_0 + v_0 t + \\frac{1}{2} a t^2",
    "chukylolxo": "T = 2\\pi \\sqrt{\\frac{m}{k}}",
    "chukyconlacdon": "T = 2\\pi \\sqrt{\\frac{l}{g}}",
    "dinhluatohm": "I = \\frac{\\mathcal{E}}{R_N + r}",
    "daodongdieuhoa": "x = A \\cos(\\omega t + \\varphi)",
    "dinhluatnewtonhai": "\\vec{F} = m\\vec{a}",
    "lucmasat": "F_{ms} = \\mu N",
    "lucdanhoi": "F_{dh} = -k \\Delta l",
    "luchapdan": "F_{hd} = G \\frac{m_1 m_2}{r^2}",
    "tongtrodien": "Z = \\sqrt{R^2 + (Z_L - Z_C)^2}",
    "dolechphadien": "\\tan \\varphi = \\frac{Z_L - Z_C}{R}",
    "congsuatdien": "P = UI \\cos \\varphi = I^2 R",
    "nangluongphoton": "\\varepsilon = hf = \\frac{hc}{\\lambda}",
    "dohutkhoi": "\\Delta m = [Z m_p + (A-Z) m_n] - m_X",
    "nangluonglienket": "W_{lk} = \\Delta m \\cdot c^2",
    "dinhluatphongxa": "N = N_0 \\cdot 2^{-\\frac{t}{T}} = N_0 e^{-\\lambda t}",
    "buocsong": "\\lambda = v \\cdot T = \\frac{v}{f}",
    "vansang": "x = k \\frac{\\lambda D}{a}",
    "vantoi": "x = \\left(k + \\frac{1}{2}\\right) \\frac{\\lambda D}{a}",
    "khoangvan": "i = \\frac{\\lambda D}{a}",
    "tuthong": "\\Phi = B \\cdot S \\cdot \\cos \\alpha",
    "suatdiendongcamung": "e_c = -\\frac{\\Delta \\Phi}{\\Delta t}",
    "luclorentz": "F = |q| v B \\sin \\alpha",

    // ==========================================
    // HÓA HỌC
    // ==========================================
    "khoiluongtrungbinh": "\\overline{M} = \\frac{#? \\cdot #? + #? \\cdot #?}{100}",
    "khilytuong": "pV = nRT",
    "tinhph": "pH = -\\log \\left[ H^+ \\right]",
    "baotoane": "\\sum n_{e(\\text{nhường})} = \\sum n_{e(\\text{nhận})}",
    "hangsocanbang": "K_c = \\frac{[#?]^#? \\cdot [#?]^#?}{[#?]^#? \\cdot [#?]^#?}",
    "dinhluatfaraday": "m = \\frac{A}{n} \\cdot \\frac{It}{F}",
    "xuctacniken": "\\xrightarrow{Ni, t^\\circ}",
    "xuctacaxit": "\\xrightarrow{H_2SO_4 \\text{ đ}, t^\\circ}",
    "nongdophantram": "C\\% = \\frac{m_{ct}}{m_{dd}} \\cdot 100\\%",
    "nongdomol": "C_M = \\frac{n}{V}",
    "khoiluongrieng": "D = \\frac{m}{V}",
    "lienkethydro": "H = 2A + 3G",

    // ==========================================
    // SINH HỌC
    // ==========================================
    "quanthehardy": "p^2 AA + 2pq Aa + q^2 aa = 1",
    "chieudaiadn": "L = \\frac{N}{2} \\times 3,4 \\text{ \\AA}",
    "tansohuhoanvigen": "f = \\frac{\\text{Số cá thể hoán vị}}{\\text{Tổng số cá thể}} \\times 100\\%",
    "dotbienlechboi": "2n \\pm 1", 
    "chukyxoan": "C = \\frac{N}{20}",
    "axitamin": "\\frac{N}{6} - 1",
  
};