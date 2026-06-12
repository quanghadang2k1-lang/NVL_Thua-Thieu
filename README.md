# NVL_Thua-Thieu
Tính NVL Thừa thiếu
# Overall Purpose

Mục tiêu chính của chương trình này là tự động hóa quy trình phức tạp **Phân bổ linh kiện tồn theo KHSX (BOM Matching & Stock Distribution)**.

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

- Script xác định mức độ phổ biến **Popularity** của từng linh kiện (một linh kiện được sử dụng bởi bao nhiêu sản phẩm khác nhau).
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
- Nếu Part A và Part B đều có thể được sử dụng cho cùng một sản phẩm, chúng sẽ được đưa vào cùng một nhóm.
- 

### Product Priority

- Người dùng sắp xếp thứ tự ưu tiên cho các sản phẩm.
- Thuật toán sẽ cố gắng đáp ứng nhu cầu vật tư của các sản phẩm có độ ưu tiên cao trước.

---

## 6. Initial Allocation (The "Greedy" Pass)

Trong mỗi pool, các component được sắp xếp theo:

1. **Popularity** (ít được chia sẻ nhất trước)
2. **Total Stock**

Thuật toán sẽ trừ dần tồn kho để đáp ứng nhu cầu của các sản phẩm theo thứ tự ưu tiên.

Nếu nhu cầu của một sản phẩm vượt quá lượng tồn kho khả dụng, phần nhu cầu chưa được đáp ứng sẽ được ép sang component **Main**, làm cho **Remaining Stock** của component đó trở thành số âm. Điều này giúp làm nổi bật các trường hợp thiếu hụt vật tư.

---

## 7. Dynamic Flaw Resolution (The "Swapping" Algorithm)

- Script quét các component có **Remaining Stock** âm.
- Nếu Product 1 bị thiếu Part X do Product 2 đã sử dụng phần tồn kho đó, thuật toán sẽ kiểm tra xem Product 2 có thể chuyển sang sử dụng một **Alternative Part Y** còn tồn kho hay không.
- Thuật toán sẽ liên tục thực hiện các thao tác **swapping** giữa các component thay thế, dịch chuyển các allocation qua lại để loại bỏ các giá trị tồn kho âm bất cứ khi nào còn tồn tại alternative part có đủ tồn kho.
- Quá trình này giúp tối ưu việc sử dụng vật tư và giảm thiểu tình trạng thiếu hụt khi có các phương án thay thế khả dụng.
