# Công cụ web rà soát thiết bị CNTT

Công cụ này phục vụ CV 2366/BĐVN-CNTT ngày 02/06/2026: nhập dữ liệu thiết bị, lấy nhanh thông tin máy tính, nhập lại file Excel đã điền và xuất báo cáo theo đúng mẫu `Ra_Soat_Thiet_Bi_CNTT.xlsx`.

## 🌐 Triển khai GitHub Pages (Miễn phí, không cần server)

**URL trực tuyến**: https://vietxuanvnp-hash.github.io/qltb_sonla/

Ứng dụng chạy hoàn toàn trên GitHub Pages (web tĩnh). Dữ liệu phiếu khảo sát được lưu trực tiếp vào file `data/inventory.json` trong repository này qua GitHub API.

Xem hướng dẫn chi tiết tại [`DEPLOY_GITHUB_PAGES.md`](DEPLOY_GITHUB_PAGES.md).

## Luồng sử dụng

1. **Nhân viên** chạy file `lay_thong_tin.bat` trên máy tính của mình.
2. Trình duyệt tự động mở URL GitHub Pages và **tự điền** hostname, CPU, RAM, ổ cứng, IP, MAC, serial.
3. Nhân viên điền thêm thông tin người sử dụng (tên, bưu cục, bộ phận) rồi nhấn **Gửi phiếu**.
4. Phiếu được lưu vào `data/inventory.json` trên GitHub (cập nhật tức thì).
5. Quản trị viên bấm **Đăng nhập quản trị**, nhập GitHub Personal Access Token để xem/sửa/xóa dữ liệu, nhập Excel và xuất báo cáo.
6. Nếu đơn vị đã có file Excel theo mẫu, quản trị viên chọn file `.xlsx` và bấm **Nhập vào kho**.
7. Bấm **Xuất Excel mẫu** để tạo file tổng hợp gửi kèm công văn.

## Cấu trúc dự án

```
device-inventory-web/
├── .github/workflows/pages.yml   ← GitHub Actions: tự deploy khi push
├── static/                       ← Thư mục web tĩnh (GitHub Pages)
│   ├── index.html
│   ├── app.js                    ← Logic ứng dụng + GitHub API
│   ├── github-config.js          ← Cấu hình repo và write token
│   ├── styles.css
│   └── vnpost-logo.png
├── data/
│   └── inventory.json            ← Kho dữ liệu phiếu khảo sát
├── assets/
│   └── Ra_Soat_Thiet_Bi_CNTT.xlsx ← Biểu mẫu gốc
├── lay_thong_tin.bat             ← Script lấy thông tin máy → mở URL
├── client_collector.ps1          ← PowerShell agent (dùng thay bat)
└── DEPLOY_GITHUB_PAGES.md        ← Hướng dẫn deploy
```

## Triển khai máy chủ nội bộ (tùy chọn)

Nếu muốn chạy server Python nội bộ (không cần internet):

```powershell
.\start_tool.ps1         # Chạy local http://127.0.0.1:8789/
.\start_intranet_server.bat  # Chạy trên LAN http://10.43.128.10:8789/
```

Xem hướng dẫn chi tiết tại [`DEPLOY_INTERNAL_SERVER.md`](DEPLOY_INTERNAL_SERVER.md).
