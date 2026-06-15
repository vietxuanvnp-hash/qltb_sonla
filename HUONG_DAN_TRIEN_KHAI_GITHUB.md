# HƯỚNG DẪN TRIỂN KHAI ỨNG DỤNG LÊN GITHUB REPOSITORY MỚI
## (Hoàn toàn Miễn phí & Tự động cấu hình)

Tài liệu này hướng dẫn chi tiết từng bước để triển khai ứng dụng kiểm kê thiết bị lên một **tài khoản GitHub mới** hoặc **Repository mới**. 

Nhờ các cập nhật tối ưu mới nhất, hệ thống sẽ **tự động cấu hình (Zero-Config)**: tự động nhận diện tên Repository của bạn và tự động gắn link tải file `lay_thong_tin.bat` tương thích hoàn toàn với website mới mà không cần chỉnh sửa code thủ công.

---

## 📌 Các bước triển khai chi tiết

### Bước 1: Tạo Repository mới trên GitHub
1. Đăng nhập vào tài khoản [GitHub](https://github.com/) của bạn.
2. Tại trang chủ GitHub, nhấn nút **New** (hoặc truy cập nhanh qua: [https://github.com/new](https://github.com/new)).
3. Cấu hình như sau:
   - **Repository name**: Nhập tên repository bất kỳ bạn muốn (Ví dụ: `quan-ly-thiet-bi` hoặc `Check_thiet_bi`).
   - **Public/Private**: Chọn **Public** (Bắt buộc chọn Public để sử dụng GitHub Pages miễn phí).
   - **Initialize this repository with**: Giữ nguyên mặc định, **KHÔNG** tích chọn thêm bất kỳ mục nào (Không thêm README, .gitignore hay License).
4. Nhấn nút **Create repository** ở dưới cùng.

---

### Bước 2: Đẩy (Push) mã nguồn lên Repository mới
1. Mở thư mục chứa code ứng dụng (`device-inventory-web`) trên máy tính của bạn.
2. Nhấp chuột phải vào một vùng trống trong thư mục và chọn **Open in Terminal** (hoặc mở **PowerShell / CMD** rồi dùng lệnh `cd` để di chuyển vào thư mục này).
3. Chạy lần lượt các lệnh Git dưới đây để liên kết và đẩy toàn bộ mã nguồn lên repository mới của bạn:

```powershell
# 1. Khởi tạo Git cục bộ (nếu chưa có)
git init

# 2. Chuyển sang nhánh chính tên là 'main'
git checkout -b main

# 3. Liên kết với Repository mới của bạn 
# (Thay thế 'TEN_GITHUB_CUA_BAN' và 'TEN_REPO_MOI' bằng thông tin thật của bạn)
git remote add origin https://github.com/TEN_GITHUB_CUA_BAN/TEN_REPO_MOI.git

# 4. Thêm toàn bộ các tệp vào hàng chờ commit
git add .

# 5. Tạo một commit đầu tiên
git commit -m "Trien khai ung dung len repo moi"

# 6. Đẩy mã nguồn lên GitHub (thêm tham số -f để đè lên các cấu hình cũ nếu có)
git push -u origin main -f
```

---

### Bước 3: Tạo Personal Access Token (Fine-Grained PAT)
Token này cho phép ứng dụng chạy trên trình duyệt của nhân viên gửi và ghi dữ liệu trực tiếp vào tệp `data/inventory.json` trong repository của bạn mà không cần máy chủ trung gian.

1. Truy cập trực tiếp trang quản lý Token của GitHub: [GitHub Personal Access Tokens (Beta)](https://github.com/settings/tokens?type=beta)
2. Nhấn nút **Generate new token**.
3. Cài đặt các thông tin như sau:
   - **Token name**: Điền tên gợi nhớ, ví dụ: `device-inventory-write-token`.
   - **Expiration**: Chọn thời hạn hết hạn cho token (ví dụ: `Custom` -> chọn tối đa 1 năm hoặc tùy chọn của bạn).
   - **Repository access**: Chọn **Only select repositories** -> Nhấp vào ô chọn và tìm kiếm, nhấp chọn đúng tên Repository bạn vừa tạo ở Bước 1.
   - **Permissions** -> Nhấp mở rộng mục **Repository permissions**:
     - Tìm dòng **Contents** -> Thay đổi quyền từ *No access* thành **Read and Write**.
4. Cuộn xuống dưới cùng và nhấn nút **Generate token**.
5. **QUAN TRỌNG**: Sao chép (Copy) đoạn mã token vừa được tạo (bắt đầu bằng chữ `github_pat_...`). **Hãy lưu mã này vào một file text trên máy tính của bạn**, vì GitHub chỉ hiển thị mã này duy nhất một lần này, nếu tải lại trang bạn sẽ không xem lại được nữa!

---

### Bước 4: Cấu hình GitHub Secret `WRITE_TOKEN`
Để ẩn mã Token bảo mật không bị lộ trong mã nguồn công khai, chúng ta sẽ đưa nó vào mục Secret của repository. GitHub Actions sẽ tự động đọc mã này khi triển khai trang web.

1. Trên trang repository của bạn trên GitHub, nhấn vào mục **Settings** (Cài đặt) ở thanh menu phía trên.
2. Ở cột menu bên trái, cuộn xuống tìm và nhấp chọn **Secrets and variables** -> Chọn tiếp **Actions**.
3. Nhấp chọn nút **New repository secret** (nút màu xanh lá ở góc phải).
4. Nhập các thông tin:
   - **Name**: Nhập chính xác tên viết hoa: `WRITE_TOKEN` (bắt buộc phải đúng chữ này).
   - **Secret**: Dán đoạn mã token `github_pat_...` bạn đã sao chép ở Bước 3.
5. Nhấn **Add secret** để lưu lại.

---

### Bước 5: Kích hoạt GitHub Pages
Chúng ta sẽ chuyển quyền xây dựng trang web cho GitHub Actions tự động thực thi.

1. Vẫn trong trang **Settings** (Cài đặt) của repository trên GitHub.
2. Ở cột menu bên trái, nhấp chọn mục **Pages**.
3. Tại phần **Build and deployment** -> Mục **Source**, nhấp chọn menu thả xuống và đổi từ *Deploy from a branch* sang **GitHub Actions**.
4. Hệ thống sẽ tự động lưu lại. Lúc này, mỗi khi bạn cập nhật code hoặc đẩy commit mới lên nhánh `main`, hệ thống sẽ tự động cập nhật lại trang web cho bạn.

---

### Bước 6: Kiểm tra tiến trình Deploy và Sử dụng
1. Nhấp chọn tab **Actions** ở menu trên cùng của Repository trên GitHub.
2. Bạn sẽ thấy một tiến trình tên là `Deploy to GitHub Pages` đang chạy (đang có biểu tượng vòng tròn xoay màu vàng).
3. Chờ khoảng 1–2 phút để tiến trình này chạy xong. Khi tất cả các bước đổi sang màu xanh lá cây có dấu tích `✓`, trang web của bạn đã được xuất bản trực tuyến thành công!
4. Nhấp vào tiến trình đã hoàn thành, bạn sẽ thấy đường link trang web hiển thị ở phần kết quả (định dạng: `https://TEN_GITHUB_CUA_BAN.github.io/TEN_REPO_MOI/`). 
5. Nhấp vào đường link đó để mở trang web kiểm kê thiết bị của bạn.

---

## ⚡ Kiểm tra hoạt động của hệ thống mới

Sau khi trang web đã hoạt động trên Repository mới, hãy thực hiện kiểm tra như sau:

1. **Tải công cụ**: Tại trang web của bạn, nhấp vào nút **"Tải công cụ lấy thông tin máy tính"**.
2. **Chạy thử**: Chạy file `lay_thong_tin.bat` vừa được tải về máy tính.
   - *Lưu ý*: Script sẽ tự động lấy thông tin phần cứng máy tính và mở trình duyệt mặc định, tự động trỏ đến đường dẫn website mới của bạn kèm dữ liệu điền sẵn trong các ô thông tin.
3. **Gửi phiếu**: Nhập tên người dùng, chọn bưu cục và nhấn **"Gửi thông tin thiết bị"**.
4. **Kiểm tra dữ liệu**: 
   - Quay lại trang repository GitHub của bạn.
   - Kiểm tra thư mục `data/` xem tệp dữ liệu `inventory.json` đã được cập nhật bản ghi mới chưa.
5. **Trang quản trị (Admin)**:
   - Trên website, nhấp vào nút **Quản trị hệ thống** ở phía trên cùng.
   - Nhập mã Token `github_pat_...` của bạn làm mật khẩu đăng nhập.
   - Bạn sẽ xem được danh sách thiết bị và có thể xuất Excel báo cáo.

---

## 💡 Lưu ý nâng cao (Nếu cần)

- **Cấu hình GPO / Agent chạy ngầm**: Nếu bạn muốn triển khai cài đặt script chạy ngầm trên máy nhân viên thông qua GPO (chính sách nhóm của Active Directory):
  - Hãy mở file `client_collector.ps1` và `install_client_protocol.ps1` trong thư mục code.
  - Sửa biến `$ServerUrl` ở đầu các file đó thành địa chỉ trang web GitHub Pages mới của bạn (ví dụ: `https://TEN_GITHUB_CUA_BAN.github.io/TEN_REPO_MOI`).
  - Thực hiện theo hướng dẫn trong file [DEPLOY_CLIENT_AGENT_GPO.md](file:///c:/Users/ADMIN/Documents/Ph%E1%BA%A7n%20m%E1%BB%81m%20l%E1%BA%A5y%20d%E1%BB%AF%20li%E1%BB%87u%20doanh%20thu%20BCCP/device-inventory-web/DEPLOY_CLIENT_AGENT_GPO.md).
