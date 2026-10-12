# BÁO CÁO PHÂN TÍCH & ĐỐI SOÁT DỮ LIỆU BÁO CÔNG SẢN XUẤT MES / ERP
**Lệnh sản xuất:** `GD260814-0002` | **Sản phẩm:** `8991CTIA13` (Mã `8-1-A-8991CTIA13-001`)

---

## 📌 PHẦN 1: PHÂN TÍCH BẢN MỚI `生产报工明细表 (3).xlsx` (297 dòng)

Qua phân tích và đối soát toàn bộ 297 dòng báo công trong file **`生产报工明细表 (3).xlsx`**, hệ thống đã phát hiện biến động quan trọng so với bản cũ và các lỗi logic tồn đọng nghiêm trọng:

### 1.1. So sánh biến động: 19 dòng bị xóa giữa Bản (1) và Bản (3)
So với bản (1) gồm 316 dòng, bản (3) đã **loại bỏ 19 dòng báo công**, bao gồm:
1. **Xóa phiếu báo nhanh KSBG:** Phiếu `KSBG-20260909-0001` (Dòng 114 ở bản cũ, 199 chiếc Tiện đa giác không mã PGD) đã bị xóa.
2. **Cắt cụt lệnh PGD-260814-0001:** Xóa sạch 8 công đoạn liên tiếp từ `CNC2` đến `成品清洗`, chỉ để lại duy nhất bước `CNC1` (198 chiếc). Khiến lệnh này biến thành lệnh dang dở / mồ côi ngay từ bước đầu.
3. **Cắt cụt lệnh PGD-260814-0003:** Xóa 4 bước sau (`无心磨2`, `CNC3`, `成品清洗`, `包装`), chỉ dừng lại ở bước `喷丸 BẮN PEENING`.
4. **Mất dòng Đóng gói (包装) tại 4 lệnh:** `PGD-0012`, `PGD-0018`, `PGD-0023` bị xóa dòng đóng gói, khiến lệnh đã làm đến Rửa/CNC3 nhưng chưa có kết quả đóng gói nhập kho.
5. **Gây ra lỗi Nhảy cóc (Bypass) tại 2 lệnh:** `PGD-0020` và `PGD-0024` bị xóa dòng `成品清洗`, nhưng dòng `包装` vẫn giữ lại, dẫn đến tình trạng chi tiết làm xong `CNC3` nhảy cóc trực tiếp sang `包装` mà không qua rửa.

> **Cập nhật cấu hình công cụ kiểm tra (MES/ERP Tool):** Theo yêu cầu, hệ thống đã **loại bỏ 2 nhóm lỗi Đơn trọng (Weight) và Ngược thời gian (Time)** khỏi bảng phát hiện lỗi và bộ lọc. Công cụ hiện tập trung tối đa vào các lỗi luồng vận hành thực tế: **Nhảy cóc (Bypass), Cắt cụt luồng (Truncation), Báo trùng (Duplicate), Báo khống số lượng (Quantity) và Thao tác dư thừa (Redundant)** (Tổng cộng 37 phát hiện).

---

### 1.2. Các bất thường cốt lõi còn tồn tại trong Bản (3)

#### ⚖️ A. Nghịch lý Đơn trọng (单重): CNC3 (78.20g) > 无心磨2 (78.18g)
- 25 bản ghi công đoạn `CNC3` đều ghi nhận Đơn trọng = **78.20 g**, cao hơn bước mài trước đó (`无心磨2` = **78.18 g**).
- Về mặt gia công cơ khí, chi tiết sau khi mài bóng đưa vào máy phay CNC3 không thể nào nặng thêm 0.02 g!

#### 🏷️ B. Bất thường tên gọi công đoạn: Có "Lần 2" nhưng không có "Lần 1"
- Toàn bộ file có 26 bản ghi mang tên `无心磨2 MÀI VÔ TÂM LẦN 2`, nhưng hoàn toàn không có `无心磨1`. Dấu hiệu copy routing từ mã khác chưa được chuẩn hóa.

#### 🔄 C. Thao tác dư thừa: Mài rung (振动研磨) & Bắn peening (喷丸)
- 20 cặp lệnh có cùng đơn trọng (78.43g), cùng sản lượng良品 100%, được báo chỉ cách nhau từ 0 đến 54 giây bởi cùng 1 nhân sự. Cần gộp thành 1 công đoạn "Xử lý bề mặt" để tiết kiệm 50% thao tác nhập liệu ERP.

#### 🔁 D. Báo trùng công đoạn (Duplicate Entries)
- Lệnh `PGD-260814-0009`: Báo 2 lần `铣翅膀` (Dòng 174: 5 cái, Dòng 221: 191 cái) và 2 lần `多边车` (Dòng 175: 9 cái, Dòng 203: 191 cái).
- Lệnh `PGD-260814-0003`: Báo 2 lần `CNC1` (Dòng 291: 186 cái, Dòng 296: 10 cái).

#### 🔀 E. Nhảy cóc công đoạn (Process Bypass)
- Lệnh `PGD-260814-0006`: Nhảy cóc bỏ qua `无心磨2` (từ Bắn bi đi thẳng sang CNC3).
- Lệnh `PGD-260814-0015`: Bỏ qua cả `CNC3` và `成品清洗`, từ Mài vô tâm nhảy thẳng sang Đóng gói.
- Lệnh `PGD-260814-0020` & `PGD-260814-0024`: Bỏ qua `成品清洗` trước khi Đóng gói.

#### ⏱️ F. Báo bù ngược thời gian (Retroactive Backfilling)
- 23 bản ghi báo công cho `CNC3` và `成品清洗` do nhân viên `VCK260834` nhập bù vào ngày 09/10/2026, trong khi hàng đã báo Đóng gói từ ngày 19/09/2026 (trễ 20 ngày).

---

## 📌 PHẦN 2: BẢNG TỔNG KẾT THEO DÕI 30 LỆNH PHÁI CÔNG (BẢN 3)

| Mã PGD | Số bước | Đánh giá trạng thái | Vấn đề logic ghi nhận |
|---|:---:|:---:|---|
| `PGD-260814-0001` | 1 | ❌ Cắt cụt | Bị xóa mất từ bước 2 đến 10, chỉ còn duy nhất CNC1 |
| `PGD-260814-0002` | 8 | ⚠️ Dang dở | Dừng tại Mài vô tâm 2, chưa có CNC3, Rửa, Đóng gói |
| `PGD-260814-0003` | 8 | ⚠️ Dang dở | Trùng 2 lần CNC1; Dừng tại Bắn bi, thiếu 4 bước cuối |
| `PGD-260814-0004` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cách nhau 18s |
| `PGD-260814-0005` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cùng ngày |
| `PGD-260814-0006` | 10 | ❌ Nhảy cóc | Bỏ qua hoàn toàn Mài vô tâm 2 |
| `PGD-260814-0007` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cách nhau 34s |
| `PGD-260814-0008` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cách nhau 19s |
| `PGD-260814-0009` | 13 | ❌ Trùng lặp | Báo trùng 2 lần Phay cánh & 2 lần Tiện đa giác |
| `PGD-260814-0010` | 11 | ⚠️ Báo bù | Đóng gói 19/09, báo bù CNC3 ngày 06/10 (sau 17 ngày) |
| `PGD-260814-0011` | 11 | ⚠️ Báo bù | Đóng gói 19/09, báo bù CNC3 & Rửa ngày 09/10 (sau 20 ngày) |
| `PGD-260814-0012` | 10 | ⚠️ Thiếu Đóng gói | Đã làm đến Rửa thành phẩm nhưng bị mất dòng Đóng gói |
| `PGD-260814-0013` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cách nhau 18s |
| `PGD-260814-0014` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Mài rung & Peening cách nhau 29s |
| `PGD-260814-0015` | 9 | ❌ Nhảy cóc | Bỏ qua CNC3 & Rửa thành phẩm, đi thẳng vào Đóng gói |
| `PGD-260814-0016` | 10 | ⚠️ Thiếu Đóng gói | Đã làm đến Rửa thành phẩm nhưng thiếu phiếu Đóng gói |
| `PGD-260814-0017` | 11 | ⚠️ Báo khống | Bước Rửa báo khống tròn 200 chiếc (Đóng gói chỉ 198) |
| `PGD-260814-0018` | 9 | ⚠️ Thiếu Đóng gói | Mất bước Rửa và Đóng gói, dừng tại CNC3 |
| `PGD-260814-0019` | 11 | ⚠️ Báo khống | Bước Rửa báo khống tròn 200 chiếc (Đóng gói chỉ 195) |
| `PGD-260814-0020` | 10 | ❌ Nhảy cóc | Mất bước Rửa, nhảy cóc từ CNC3 sang Đóng gói |
| `PGD-260814-0021` | 11 | ⚠️ Báo bù | Đóng gói 19/09, báo bù CNC3 & Rửa ngày 09/10 (sau 20 ngày) |
| `PGD-260814-0022` | 11 | ⚠️ Có lỗi | Đơn trọng CNC3 ngược logic; Báo bù sau Đóng gói |
| `PGD-260814-0023` | 10 | ⚠️ Thiếu Đóng gói | Đã làm đến Rửa thành phẩm nhưng mất dòng Đóng gói |
| `PGD-260814-0024` | 10 | ❌ Nhảy cóc | Mất bước Rửa, nhảy cóc từ CNC3 sang Đóng gói |
| `PGD-260814-0025` | 11 | ⚠️ Báo khống | Bước Rửa báo khống tròn 200 chiếc (Đóng gói chỉ 197) |
| `PGD-260814-0026` | 11 | ⚠️ Báo bù | Đóng gói 02/10, báo bù CNC3 & Rửa ngày 09/10 (sau 7 ngày) |
| `PGD-260814-0027` | 11 | ⚠️ Báo bù | Đóng gói 19/09, báo bù CNC3 & Rửa ngày 09/10 (sau 20 ngày) |
| `PGD-260814-0028` | 11 | ⚠️ Báo bù | Đóng gói 19/09, báo bù CNC3 ngày 08/10 (sau 19 ngày) |
| `PGD-260814-0029` | 3 | ❌ Cắt cụt | Chỉ có CNC1 -> CNC2 -> Phay cánh, dừng hẳn từ bước 4 |
| `PGD-260814-0030` | 10 | ⚠️ Thiếu Đóng gói | Đã làm đến Rửa thành phẩm nhưng chưa có phiếu Đóng gói |

---

## 💡 KHUYẾN NGHỊ ĐIỀU CHỈNH & CHUẨN HÓA

1. **Chuẩn hóa Routing kỹ thuật trong hệ thống MES/ERP:**
   - Xem lại bản vẽ và công nghệ: Nếu gia công CNC xong mới mài hoàn thiện, phải đổi thứ tự thành: `CNC3 (78.20g) -> 无心磨 (78.18g)`.
   - Bỏ số "2" trong tên công đoạn `无心磨2`.
   - Nếu thực tế không cần chạy CNC3, loại bỏ hẳn CNC3 ra khỏi lưu trình chuẩn của mã `8991CTIA13`.
2. **Gộp công đoạn Xử lý bề mặt:**
   - Gộp `振动研磨` và `喷丸` thành 1 công đoạn duy nhất trên hệ thống để tiết kiệm thời gian báo công và tránh dữ liệu trùng lặp.
3. **Cài đặt ràng buộc Validation trên phần mềm:**
   - Không cho phép báo đóng gói khi các công đoạn trước chưa hoàn thành.
   - Không cho phép nhập ngày báo công lùi về quá khứ quá 24h so với công đoạn liền kề.
   - Ngăn chặn báo trùng 2 lần cho cùng một công đoạn trong 1 mã lệnh phái công.