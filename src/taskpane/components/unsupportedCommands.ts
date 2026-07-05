// Danh sách các lệnh LaTeX không thể hoặc hay bị lỗi khi dịch sang MathML của Word
export const unsupportedMathMlCommands = [
  "\\overline{",  // Lệnh gạch ngang trên
  "\\overgroup{",
  "\\overbrace{",
  "\\underbrace{",
  "\\underline{",
  "\\xrightarrow",
  "\\overrightarrow{",
  "\\phase{",
  "\\overparen{",
  "\\overleftarrow{",
  "\\overleftrightarrow{",
  "\\ce{",
];