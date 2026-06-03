Trong file JSON được tạo từ script "1_generate_json_with_textcontent.py", hãy tiền xử lý các trường title, lý_do_mục_đích, nơi_giao_dịch, loại_chứng_khoán, và text_content bằng cách loại bỏ dấu Tiếng Việt, lower key, loại bỏ khoảng trắng dư thừa và các ký tự lỗi có thể phát sinh. Sau đó, dùng chúng để phái sinh ra 1 file Excel với các trường dữ liệu được mô tả dưới đây:

* date: giữ nguyên.

* collected_date: giữ nguyên.

* url: giữ nguyên.

* text_content: giữ nguyên.

* Trường [MaChungKhoan]: là giá trị của cột code.

* Trường [LoaiQuyen], xét lần lượt các trường hợp:
  * Trả về "Tin huỷ": nếu giá trị tại cột title (được tiền xử lý) có keyword "huy" ngay sau ký tự ": " hoặc ":" đầu tiên
  * Trả về "Thay đổi": nếu giá trị tại cột title (được tiền xử lý) chứa một trong số các keywords ["thay doi", "chuyen du lieu", "chuyen san", "dinh chinh", "dieu chinh"]
  * Trả về "Đăng ký lưu ký": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["dang ky chung chi", "dang ky chung khoan", "dang ky co phieu", "dang ky trai phieu", "luu ky chung chi", "luu ky chung khoan", "luu ky co phieu", "luu ky trai phieu"]
  * Trả về "Cổ phiếu thưởng": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["phat hanh co phieu", "nhan co phieu", "nhan them co phieu"]; hoặc trong giá trị của một trong 2 trường title hoặc lý_do_mục_đích (được tiền xử lý) xuất hiện đồng thời 2 keywords "co phieu" và "de tang von"
  * Trả về "Cổ tức bằng cổ phiếu": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa keyword ["co tuc bang co phieu", "co tuc co phieu", "chuyen doi thanh co phan", "co phieu de tra co tuc"] hoặc đồng thời chứa 2 keywords "co tuc" và "co phieu"
  * Trả về "Cổ tức bằng tiền": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số những keywords ["co tuc bang tien", "co tuc tien", "thanh toan co tuc",  "tam ung co tuc", "lai trai phieu", "chi co tuc", "mua lai trai phieu", "tra goc", "thanh toan goc", "tra lai", "thanh toan lai"], hoặc đồng thời xuất hiện 2 keywords "co tuc" và "bang tien", hoặc đồng thời xuất hiện 2 keywords "mua lai" và "trai phieu", hoặc đồng thời xuất hiện 2 keywords "mua lai" và "truoc han"
  * Trả về "Quyền biểu quyết": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["y kien co dong", "dai hoi co dong", "dai hoi dong co dong", "dai hoi nha dau tu", "bang van ban", "lay y kien", "thong qua phuong an", "dai hoi", "dong co dong"]
  * Trả về "Quyền mua": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["quyen mua", "quyen mua co phieu", "quyen mua trai phieu"]
  * Trả về "Hoán đổi chuyển đổi": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["hoan doi co phieu", "chuyen doi trai phieu", "chuyen doi co phieu", "chuyen quyen"]
  * Trả về "Khai báo chứng quyền": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa đồng thời keywords "chung quyen" và "dang ky", hoặc đồng thời "chung quyen" và "giay chung nhan", hoặc đồng thời "chung quyen" và "khai bao"
  * Trả về "Thực hiện chứng quyền": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa keywords "thuc hien chung quyen"
  * Trả về null: nếu không thoả bất kỳ trường hợp nào vừa nêu.

* Trường [ThiTruong], xét lần lượt các trường hợp:
  * Nếu trường nơi_giao_dịch (được tiền xử lý) không null, lần lượt xét:
    * Trả về "Upcom": nếu nơi_giao_dịch (được tiền xử lý) chứa keyword "upcom"
    * Trả về "Niêm yết": nếu nơi_giao_dịch (được tiền xử lý) có chứa một trong các keywords ["hnx", "hose"]
    * Trả về "Trái phiếu": nếu nơi_giao_dịch (được tiền xử lý) có chứa keyword "trai phieu"
  * Nếu trường nơi_giao_dịch (được tiền xử lý) là null hoặc chuỗi rỗng, lần lượt xét:
    * Trả về "UpCom": khi trường title hoặc text_content (được tiền xử lý) có chứa keyword "upcom"
    * Trả về "Trái phiếu": nếu một trong số các trường loại_chứng_khoán, title, hoặc text_content (được tiền xử lý) có chứa keyword "trai phieu"
    * Trả về "Niêm yết": khi một trong số các trường loại_chứng_khoán, title, hoặc text_content (được tiền xử lý) có chứa keywords "co phieu"
  * Trả về null: nếu cả nơi_giao_dịch và loại_chứng_khoán và title và text_content đều null; hoặc không rơi vào bất kỳ trường hợp nào nêu trên.

* Trường [SanGD], xét lần lượt các trường hợp:
  * Trả về "HSX": nếu nơi_giao_dịch (được tiền xử lý) là "hose"
  * Trả về "HNX": các trường hợp còn lại

* Trường [NoiQuanLy], xét lần lượt các trường hợp:
  * Trả về "Chi nhánh VSD": nếu text_content có chứa keywords "cnvsdc"
  * Trả về "Hội sở VSD": nếu text_content có chứa keywords "vsdc"
  * Trả về null: nếu không rơi vào 2 trường hợp trên.

* Trường [NgayChot]:
  * Nếu giá trị trường ngày_đăng_ký_cuối không null -> Trả ra giá trị của trường ngày_đăng_ký_cuối
  * Nếu giá trị trường ngày_đăng_ký_cuối là null -> Tìm keyword "ngay dang ky cuoi" hoặc "thoi gian dang ky cuoi" hoặc "ngay chot" hoặc "thoi gian chot" đầu tiên xuất hiện, xét 200 ký tự tiếp theo và lấy ra chuỗi ký tự sớm nhất tìm được có định dạng ngày. Chuỗi ký tự chứa thông tin ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó.

* Trường [NgayThucHien]:
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là một trong số các loại ["Quyền biểu quyết", "Cổ phiếu thưởng", "Cổ tức bằng cổ phiếu"] -> Trong text_content (được tiền xử lý), tìm chuỗi ký tự "thoi gian thuc hien" hoặc "ngay thuc hien" hoặc "thoi gian hien" hoặc "ngay hien" hoặc "thoi gian thuc" hoặc "ngay thuc", sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra chuỗi ký tự có định dạng ngày đầu tiên tìm được, tuy nhiên, nếu tìm thấy 2 chuỗi ký tự có định dạng ngày thì lấy chuỗi ở gần cuối hơn. Chuỗi ký tự chứa thông tin ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó.
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là các loại còn lại, hoặc trong text_content không tìm thấy các keyword-> trả về null.

* Trường [NgayThanhToan]:
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là một trong số các loại ["Quyền biểu quyết", "Cổ phiếu thưởng", "Cổ tức bằng cổ phiếu"] -> trả về null
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là các loại còn lại -> Trong text_content (được tiền xử lý), tìm một trong số các keywords ["ngay thuc hien", "ngay thanh toan", "thoi gian thuc hien", "thoi gian thanh toan"], sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra chuỗi ký tự có định dạng ngày đầu tiên tìm được, tuy nhiên, nếu tìm thấy 2 chuỗi ký tự có định dạng ngày thì lấy chuỗi ở gần cuối hơn. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó.

* Trường [NgayHetHanCNQuyenMua]:
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là "Quyền mua" -> Trong text_content (được tiền xử lý), xét dòng đầu tiên có chứa keyword "thoi gian chuyen nhuong" hoặc "ngay chuyen nhuong" hoặc "han chuyen nhuong", trả về chuỗi ký tự có định dạng ngày ở gần cuối nhất của dòng. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là các loại còn lại -> trả về null.

* Trường [NgayHetHanDKQuyenMua]:
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là "Quyền mua" -> Trong text_content (được tiền xử lý), xét dòng đầu tiên có chứa keyword "thoi gian dang ky" hoặc "ngay dang ky" hoặc "han dang ky", trả về chuỗi ký tự có định dạng ngày ở gần cuối nhất của dòng. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó
  * Nếu giá trị tại trường [LoaiQuyen] (như rule bên trên) là các loại còn lại -> trả về null.

* Trường [DonViHuongQuyen] & trường [GiaTriHuongQuyen]:
  * Đơn vị hưởng quyền (DonViHuongQuyen) là số cổ phiếu nắm giữ đơn vị (mẫu số của tỷ lệ thực hiện) dùng để xác định một giá trị hưởng quyền tương ứng được quy định (GiaTriHuongQuyen, tử số của tỷ lệ thực hiện). Ví dụ, với "Tỷ lệ thực hiện: 50%/cổ phiếu (1 cổ phiếu được nhận 5.000 đồng)" thì Đơn vị hưởng quyền là 1 và Giá trị hưởng quyền là 5000
  * Phạm vi tìm kiếm 2 thông tin [DonViHuongQuyen] & [GiaTriHuongQuyen] là trong text_content từ sau các keywords "ty le thuc hien" hoặc "ty le thanh toan" cho đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước)
  * Giả sử Đơn vị hưởng quyền là X & Giá trị hưởng quyền là Y thì chúng có thể được tìm thấy từ 1 trong số những mẫu nội dung sau: "<X> cổ phiếu được nhận <Y>", "<X> cổ phiếu nhận được <Y>", "<X> cổ phiếu được <Y>", "<X> cổ phiếu nhận <Y>", "<X> trái phiếu được nhận <Y>", "<X> trái phiếu nhận được <Y>", "<X> trái phiếu được <Y>", "<X> trái phiếu nhận <Y>", "<X> cổ phiếu sẽ được nhận <Y>", "<X> cổ phiếu sẽ nhận được <Y>", "<X> cổ phiếu sẽ được <Y>", "<X> cổ phiếu sẽ nhận <Y>", "<X> trái phiếu sẽ được nhận <Y>", "<X> trái phiếu sẽ nhận được <Y>", "<X> trái phiếu sẽ được <Y>", "<X> trái phiếu sẽ nhận <Y>", "<X> cổ phiếu sẽ được nhận thêm <Y>", "<X> cổ phiếu sẽ nhận được thêm <Y>", "<X> cổ phiếu sẽ được thêm <Y>", "<X> cổ phiếu sẽ nhận thêm <Y>", "<X> trái phiếu sẽ được nhận thêm <Y>", "<X> trái phiếu sẽ nhận được thêm <Y>", "<X> trái phiếu sẽ được thêm <Y>", "<X> trái phiếu sẽ nhận thêm <Y>", "<X> cổ phiếu được nhận thêm <Y>", "<X> cổ phiếu nhận được thêm <Y>", "<X> cổ phiếu được thêm <Y>", "<X> cổ phiếu nhận thêm <Y>", "<X> trái phiếu được nhận thêm <Y>", "<X> trái phiếu nhận được thêm <Y>", "<X> trái phiếu được thêm <Y>", "<X> trái phiếu nhận thêm <Y>", "<X> cổ phiếu - <Y> quyền biểu quyết", "<X> trái phiếu - <Y> quyền biểu quyết", "<X> cổ phiếu – <Y> quyền biểu quyết", "<X> trái phiếu – <Y> quyền biểu quyết", "<X> cổ phiếu — <Y> quyền biểu quyết", "<X> trái phiếu — <Y> quyền biểu quyết", "<X> chứng chỉ quỹ - <Y> quyền biểu quyết", "<X> chứng chỉ quỹ – <Y> quyền biểu quyết", "<X> chứng chỉ quỹ — <Y> quyền biểu quyết", "<X>:<Y>", "<X> : <Y>", "<X>/<Y>", "<X> / <Y>". Ngoài ra, trong những mẫu nội dung vừa nêu, có thể có những trường hợp có thêm mô tả bằng chữ ngay sau X theo dạng "<X> (<mô tả bằng chữ của X>)<...>" thì cần bắt được X chuẩn xác; ví dụ, trường hợp "01 (một) trái phiếu nhận được 1.051.000.000" thì cần xác định được X=1.
  * Đặc biệt lưu ý, nếu Giá trị hưởng quyền sau khi trích xuất là một số thập phân với dấu "," thì cần có thêm 1 bước xử lý để ra được kết quả cuối, đó là nhân Đơn vị hưởng quyền và Giá trị hưởng quyền với cùng một số luỹ thừa của 10 để Giá trị hưởng quyền là một số vừa nguyên và tỷ lệ thực hiện (Giá trị hưởng quyền/ Đơn vị hưởng quyền) không thay đổi so với số gốc. Ví dụ, với trường hợp nội dung "100 cổ phiếu được nhận 73.500,7476" thì thông tin gốc lấy được là Đơn vị hưởng quyền là 100 & Giá trị hưởng quyền là 73500,7476, tuy nhiên, 73500,7476 là một số thập phân với 4 chữ số sau dấu phẩy nên ta cần nhân cả Đơn vị hưởng quyền & Giá trị hưởng quyền cho 1e4 để ra giá trị cuối cùng của Đơn vị hưởng quyền là 1000000 & Giá trị hưởng quyền là 735007476.

* Trường [NoiDung]:
  * Nếu trường [LoaiQuyen] (như rule bên trên) là một trong các loại ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký lưu ký", "Thay đổi"] -> trả ra null
  * Nếu trường [LoaiQuyen] (như rule bên trên) là "Tin huỷ", xét nếu một trong các trường title hoặc text_content (được tiền xử lý):
    * có chứa keyword "huy dang ky chung khoan" -> trả ra "Hủy đăng ký chứng khoán"
    * có chứa keyword "huy dang ky chung quyen" -> trả ra "Hủy đăng ký chứng quyền"
    * có chứa keyword "huy dang ky trai phieu" -> trả ra "Hủy đăng ký trái phiếu"
    * có chứa keyword "huy dot chot danh sach thuc hien chung quyen" -> trả ra "Hủy đợt chốt danh sách thực hiện chứng quyền"
    * có chứa keyword "huy danh sach nguoi so huu chung khoan" -> trả ra "Hủy danh sách người sở hữu chứng khoán"
    * có chứa keyword "huy thong bao ngay dang ky cuoi cung" -> trả ra "Hủy thông báo ngày đăng ký cuối cùng"
  * Nếu trường [LoaiQuyen] (như rule bên trên) là các loại còn lại -> Lấy giá trị của cột title mà đã được loại bỏ giá trị của cột code và ký tự ": " (ví dụ, nếu title là "ACB: Chi trả cổ tức bằng tiền năm 2025" thì sẽ trả ra giá trị "Chi trả cổ tức bằng tiền năm 2025"). Sau đó, nối thêm vào nó 3 thông tin được lấy trong text_content ở định dạng gốc có dấu Tiếng Việt (trước khi tiền xử lý) và nối với nhau bởi dấu " - " gồm: một là Tỷ lệ thực hiện (hoặc Tỷ lệ thanh toán), hai là Ngày thanh toán (hoặc Ngày thực hiện, Thời gian thực hiện, Thời gian thanh toán), và ba là Giá phát hành.  Ví dụ: code THS có [LoaiQuyen]="Cổ tức bằng tiền" thì Giá trị của cột [NoiDung] sẽ là "Chi trả cổ tức bằng tiền năm 2025 - Tỷ lệ thực hiện: 8%/ cổ phiếu (1 cổ phiếu được nhận 800 đồng) - Ngày thanh toán: 15/05/2026 - Giá phát hành: 10.000 đồng". Trong text_content (được tiền xử lý), thông tin Tỷ lệ thực hiện được bắt đầu bằng keyword "ty le thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Tỷ lệ thanh toán được bắt đầu bằng "ty le thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng); Ngày thực hiện được bắt đầu bằng "ngay thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Ngày thanh toán được bắt đầu bằng "ngay thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Thời gian thực hiện được bắt đầu bằng "thoi gian thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Thời gian thanh toán được bắt đầu bằng "thoi gian thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng); Giá phát hành được bắt đầu bằng "gia phat hanh" và lấy đến hết dòng (khi gặp ký tự xuống dòng).

* Trường [is_completed] là một trường boolean trả ra giá trị 0 hoặc 1:
  * Trả ra giá trị 1 khi thoả mãn cả 4 tiêu chí sau:
    * Cả 6 trường [LoaiQuyen], [MaChungKhoan], [ThiTruong], [SanGD], [NoiQuanLy], [NgayChot] đều không null
    * Nếu [LoaiQuyen] là "Quyền mua" -> cả 2 trường [NgayHetHanDKQuyenMua] & [NgayHetHanCNQuyenMua] đồng thời không null
    * Nếu [LoaiQuyen] là "Quyền biểu quyết" -> trường [NgayThucHien] không null
    * Nếu [LoaiQuyen] là "Cổ tức bằng tiền" -> trường [NgayThanhToan] không null.
  * Còn lại, trả ra giá trị 0.

* Trường [is_special] là một trường boolean trả ra giá trị 0 hoặc 1:
  * Trả ra 1 khi [LoaiQuyen] là một trong các loại ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký lưu ký", "Tin huỷ", "Thay đổi"]
  * Còn lại, trả ra giá trị 0.

