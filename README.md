# LSTS Equation Editor (Easy Efficient Equation Editor)

<p align="center">
  <b>Add-in Microsoft Word thông minh dành cho giáo viên và học sinh toán học</b>
</p>

---

## 🌟 Giới thiệu dự án

**LSTS Equation Editor** là một Office Add-in dành cho Microsoft Word được phát triển bằng **React, TypeScript và Fluent UI**, tích hợp các thư viện toán học mạnh mẽ như **MathLive** và **CortexJS Compute Engine**. 

Dự án ra đời nhằm giải quyết triệt để các bài toán nhập liệu công thức toán học phức tạp, tính toán nhanh giá trị biểu thức, vẽ phép tính dọc (cộng, trừ, nhân, chia) và tự động tạo **Bảng biến thiên (Variation Table)** dưới dạng ảnh vector SVG sắc nét trực tiếp trong Microsoft Word.

---

## ✨ Các tính năng nổi bật

- 🧮 **Soạn thảo toán học trực quan**: Tích hợp bàn phím ảo MathLive, hỗ trợ gõ nhanh bằng phím tắt (Inline) và lệnh Macro (`\`).
- 🔄 **Đọc ngược từ Word (Smart Read)**: Hỗ trợ nhận diện và tải lại công thức từ Word lên Add-in thông qua tính năng Double-Click.
- 📐 **Vẽ phép tính dọc tự động**: Tự động nhận diện phép cộng, trừ, nhân, chia số nguyên hoặc phân số và xuất ra sơ đồ SVG chuẩn chỉnh.
- 📉 **Tự động sinh Bảng biến thiên (SVG)**: 
  - Tự động tính đạo hàm $y'$.
  - Dò nghiệm đạo hàm bằng phương pháp số học chính xác tuyệt đối.
  - Tự động xác định các điểm cực trị, giới hạn vô cực ($\pm\infty$) và vẽ mũi tên xu hướng trực quan.
- 🗂️ **Quản lý từ điển cá nhân**: Cho phép lưu trữ, xuất/nhập (`Export/Import`) các phím tắt công thức cá nhân dưới định dạng JSON.
- 🌐 **Hỗ trợ 2 ngôn ngữ**: Giao diện song ngữ Tiếng Việt và Tiếng Anh.

---

## 🛠️ Công nghệ sử dụng

- **Frontend**: React, TypeScript, Fluent UI (Office UI Fabric)
- **Math Engine**: MathLive, CortexJS Compute Engine
- **Office Integration**: Office JavaScript API (Word Web Add-in)
- **Rendering**: SVG Vector Graphics, MathJax (cho việc xử lý ký hiệu)

---

## 🚀 Hướng dẫn Cài đặt & Chạy dự án (Local Development)

### 1. Yêu cầu hệ thống
Trước khi bắt đầu, hãy đảm bảo máy tính của bạn đã cài đặt:
- **Node.js** (Phiên bản khuyến nghị: LTS 18.x trở lên)
- **npm** hoặc **yarn**
- **Microsoft Word** (Phiên bản trên Desktop chạy trên Windows hoặc macOS)

### 2. Các bước cài đặt

Clone repository này về máy và cài đặt các gói thư viện phụ thuộc:

```bash
# Clone dự án về máy
git clone https://github.com/TEN_GITHUB_CUA_ANH/EquationType.git

# Di chuyển vào thư mục dự án
cd EquationType

# Cài đặt các thư viện
npm install
```

### 3. Chạy môi trường phát triển (Dev Server)

Để chạy Add-in ở chế độ phát triển và kết nối trực tiếp với Microsoft Word trên máy tính:

```bash
npm run dev-server
```

Hoặc để khởi chạy đồng thời server và tự động bật Microsoft Word:

```bash
npm run start
```

---

## 📖 Hướng dẫn Sử dụng

1. **Mở Add-in trong Word**: Khi Word khởi động kèm theo Add-in, chọn tab **Home** và nhấn vào biểu tượng của Add-in trên thanh Ribbon để mở bảng điều khiển (Taskpane).
2. **Soạn thảo công thức**: Nhập trực tiếp biểu thức vào ô MathLive hoặc sử dụng bảng ký hiệu toán học có sẵn.
3. **Tính toán & Xử lý hàm số**:
   - Nhập hàm số dạng `y = x^2 - 3x + 2` hoặc `f(x) = x^3 - 3x`.
   - Nhấn nút **Solve**, **Evaluate** hoặc xem trực tiếp phần **Preview Variation Table** (Bảng biến thiên) ở phía dưới.
4. **Chèn vào Word**: Nhấn đúp chuột (`Double-Click`) vào khu vực xem trước hoặc nhấn nút **Insert / Update** để đưa công thức/bảng biến thiên vào vị trí con trỏ trong văn bản Word.

---

## 📂 Cấu trúc Thư mục

```text
EquationType/
├── src/
│   ├── taskpane/
│   │   ├── components/
│   │   │   ├── App.tsx             # Giao diện chính và logic điều phối
│   │   │   ├── MathPalette.tsx     # Bảng ký hiệu toán học
│   │   │   ├── mathShortcuts.ts    # Phím tắt gõ nhanh
│   │   │   └── unsupportedCommands.ts
│   │   └── index.tsx
│   └── assets/                     # Icons và tài nguyên tĩnh
├── manifest.xml                    # Cấu hình Office Add-in
├── package.json
└── README.md
```

---

## 👨‍💻 Người phát triển

- **Huỳnh Quốc Thịnh** 
- Trường THCS và THPT Đinh Thiện Lý

---

## 📄 Giấy phép (License)

Dự án này được phát hành dưới điều khoản của **MIT License**. Xem file [LICENSE](LICENSE) để biết thêm chi tiết.

```text
MIT License

Copyright (c) 2026 Huỳnh Quốc Thịnh

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```