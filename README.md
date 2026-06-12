# Mục đích

Mục tiêu chính của chương trình này là tự động hóa quy trình phức tạp **Phân bổ linh kiện tồn theo KHSX**.

Nhận các file **Bill of Materials (BOM)** cho nhiều sản phẩm khác nhau, nhân với sản lượng kế hoạch (**Kế hoạch sản xuất**), đối chiếu với tồn kho hiện có từ nhiều kho, sau đó phân bổ tồn kho cho từng sản phẩm một cách thông minh dựa trên mức độ ưu tiên do người dùng thiết lập và khả năng thay thế giữa các linh kiện.

# Logic từng bước

## 1. Nhập dữ liệu BOMs

### RDBOM & MANBOM Uploads
Đọc các file **RDBOM** và **MANBOM**.

### Làm sạch dữ liệu
Chuẩn hóa tên sản phẩm, loại bỏ các linh kiện bị trùng trong cùng một level, sau đó gộp dữ liệu từ các BOM để tính toán **Sản lượng tiêu hao (Standard Quantity)** cần cho mỗi sản phẩm (*Quantity per Product + Consumption/Scrap Rate*).

---

## 2. Phân loại linh kiện

### Popularity & Level Groups

- Xác định mức độ phổ biến **Popularity** của từng linh kiện (một linh kiện được sử dụng bởi bao nhiêu sản phẩm khác nhau).
- Tạo **Level Group**, thể hiện những sản phẩm nào và level nào đang sử dụng linh kiện đó.

### Pivot Table

Dữ liệu được chuyển thành **Pivot Table**, trong đó:

- Mỗi **Filter VNPT MAN P/N** là một dòng.
- Số lượng yêu cầu của từng sản phẩm được thể hiện dưới dạng các cột.

---

## 3. Tính sản lượng cần sản xuất theo kế hoạch

- Người dùng nhập sản lượng kế hoạch cho từng sản phẩm.
- Tính toán tổng số lượng tuyệt đối của từng linh kiện cần sử dụng (**SL theo KH**).

---

## 4. Nhập dữ liệu tồn từ các bộ phận

- Script đọc dữ liệu tồn kho từ tối đa 5 nguồn khác nhau:
  - Kho tốt
  - Kho CLC
  - Nhà máy Tech
  - Nhà máy SCBH
  - KHHV

  hoặc từ một file tồn kho đã được tổng hợp sẵn.

- Dữ liệu được gộp theo mã linh kiện để tính **Tổng tồn** trên toàn bộ các kho.

---

## 5. Nhóm và thứ tự ưu tiên

### Allocation Pools (Nhóm phân bổ)

- Thuật toán **Union-Find** được sử dụng để nhóm các linh kiện có thể thay thế cho nhau thành các **Allocation Pools**.
- Nếu Linh kiện A và Linh Kiện B đều có thể được sử dụng cho cùng một sản phẩm, chúng sẽ được đưa vào cùng một nhóm.
- Nếu Linh kiện A và B được sử dụng cho cùng 1 sản phẩm, linh kiện B và C lại có thể được dùng chung cho 1 sản phẩm khác => Cả A,B và C sẽ đều nằm trong cùng 1 nhóm phân bổ

### Thứ tự ưu tiên

- Người dùng sắp xếp thứ tự ưu tiên cho các sản phẩm. (Cân nhắc để tự động sắp xếp ưu tiên hoặc bỏ bước này)
- Thuật toán sẽ cố gắng đáp ứng nhu cầu vật tư của các sản phẩm có độ ưu tiên cao trước.

---

## 6. Phân bổ tồn bồ phận

Trong mỗi allocation pool, các linh kiện được sắp xếp theo:

1. **Popularity** (ít phổ biến trước)
2. **Total Stock** (Có tổng tồn)

Thuật toán sẽ trừ dần từ **Tổng tồn** để đáp ứng nhu cầu của các sản phẩm theo thứ tự ưu tiên.

- Nếu nhu cầu của một sản phẩm vượt quá lượng tồn kho khả dụng, phần nhu cầu chưa được đáp ứng sẽ được ép sang linh kiện **Main (linh kiện đầu tiên)**, làm cho **Remaining Stock** của linh kiện đó trở thành số âm. Điều này giúp làm nổi bật các trường hợp thiếu hụt vật tư.
- Tuy nhiên sau khi ra kết quả, sẽ có các trường hợp linh kiện tồn chưa được phân bổ một cách tối ưu nhất: Ví dụ như trong cùng 1 allocation pool, thuật toán phân bổ theo thứ tự ưu tiên thay vì lấp đầy cho sản phẩm ít nhu cầu hơn trước, dẫn đến việc có thể bị sót lại nhiều hơn 1 **Remaining Stock** bị âm.
VD: Trong 1 pool, tồn có 10,000; linh kiện cho sản phẩm A cần 14000, linh kiện cho sản phẩm B cần 3000. Nếu thứ tự ưu tiên A cao hơn B, thuật toán sẽ phân bổ hết 10,000 tồn cho A => kết quả còn lại là A bị thiếu 4000, B thiếu 3000 (SAI).
- Do đó, ta muốn thuật toán phân bổ cho B trước bất chấp thứ tự ưu tiên rồi mới phân bổ cho A => B sẽ có đủ, A chỉ thiếu 7000 (lúc SC đi mua NVL sẽ dễ dàng hơn) 

---
 
## 7. Giải quyết vấn đề phân bổ (Thuật toán hoán đổi)

- Quét các linh kiện có **Remaining Stock** âm.
- Nếu sản phẩm 1 bị thiếu linh kiện X do sản phẩm 2 đã sử dụng phần tồn kho đó, thuật toán sẽ kiểm tra xem sản phẩm 2 có thể chuyển sang sử dụng một **Linh kiện thay thế Y** còn tồn kho hay không.
- Thuật toán sẽ liên tục thực hiện các thao tác **hoán đổi** giữa các linh kiện thay thế, dịch chuyển các phân bổ qua lại để loại bỏ các giá trị tồn kho âm bất cứ khi nào còn tồn tại linh kiện thay thế có đủ tồn kho.
- Cuối cùng, đẩy linh kiện đang có **Remaining Stock** bị âm sang linh kiện thay thế mà có độ phổ biến **Popularity** lớn nhất. 
- Quá trình này giải quyết vấn đề ở bước 6, giúp tối ưu việc sử dụng vật tư và giảm thiểu tình trạng thiếu hụt khi có các phương án thay thế khả dụng.
