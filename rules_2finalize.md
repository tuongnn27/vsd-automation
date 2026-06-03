Từ các trường thông tin craw được ban đầu, hãy tiền xử lý các trường title, lý_do_mục_đích, nơi_giao_dịch, loại_chứng_khoán, và text_content bằng cách loại bỏ dấu Tiếng Việt, lower key, loại bỏ khoảng trắng dư thừa và các ký tự lỗi có thể phát sinh. Sau đó, dùng chúng để phái sinh ra 1 file Excel với các trường dữ liệu được mô tả dưới đây:

* published_at: giữ nguyên.

* collected_at: giữ nguyên.

* url: giữ nguyên.

* text_content: giữ nguyên.

* Trường [MaChungKhoan]: là giá trị của cột code.

* Trường [TieuDe]: Lấy giá trị của cột title mà đã được loại bỏ giá trị của cột code và ký tự ": " (ví dụ, nếu title là "ACB: Chi trả cổ tức bằng tiền năm 2025" thì sẽ trả ra giá trị "Chi trả cổ tức bằng tiền năm 2025")

* Trường [NhomQuyen], xét lần lượt các trường hợp:
  * Trả về "Tin huỷ": nếu giá trị tại cột title (được tiền xử lý) có keyword "huy" ngay sau ký tự ": " hoặc ":" đầu tiên, và đồng thời không chứa keyword "chung quyen"
  * Trả về "Thay đổi": nếu giá trị tại cột title (được tiền xử lý) chứa một trong số các keywords ["thay doi", "chuyen du lieu", "chuyen san", "dinh chinh", "dieu chinh"], và đồng thời không chứa keyword "chung quyen"
  * Trả về "Đăng ký Lưu ký": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["dang ky chung chi", "dang ky chung khoan", "dang ky co phieu", "dang ky trai phieu", "luu ky chung chi", "luu ky chung khoan", "luu ky co phieu", "luu ky trai phieu"], đồng thời không chứa keyword "to chuc" ngay trước nó, nghĩa là không chứa bất kỳ keywords nào trong số ["to chuc dang ky chung chi", "to chuc dang ky chung khoan", "to chuc dang ky co phieu", "to chuc dang ky trai phieu", "to chuc luu ky chung chi", "to chuc luu ky chung khoan", "to chuc luu ky co phieu", "to chuc luu ky trai phieu"]
  * Trả về "Cổ tức cổ phiếu / Cổ phiếu thưởng": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["phat hanh co phieu", "nhan co phieu", "nhan them co phieu", "co tuc bang co phieu", "co tuc co phieu", "chuyen doi thanh co phan", "co phieu de tra co tuc"]; hoặc trong giá trị của một trong 2 trường title hoặc lý_do_mục_đích (được tiền xử lý) đồng thời chứa 2 keywords "co phieu" và "de tang von", hoặc đồng thời chứa 2 keywords "co tuc" và "co phieu".
  * Trả về "Cổ tức tiền": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số những keywords ["co tuc bang tien", "co tuc tien", "thanh toan co tuc",  "tam ung co tuc", "lai trai phieu", "chi co tuc", "mua lai trai phieu", "tra goc", "thanh toan goc", "tra lai", "thanh toan lai"], hoặc đồng thời xuất hiện 2 keywords "co tuc" và "bang tien", hoặc đồng thời xuất hiện 2 keywords "mua lai" và "trai phieu", hoặc đồng thời xuất hiện 2 keywords "mua lai" và "truoc han"
  * Trả về "Quyền biểu quyết": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["y kien co dong", "dai hoi co dong", "dai hoi dong co dong", "dai hoi nha dau tu", "bang van ban", "lay y kien", "thong qua phuong an", "dai hoi", "dong co dong", "quyen de cu"]
  * Trả về "Quyền mua": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa keyword "quyen mua"
  * Trả về "Hoán đổi chuyển đổi": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["hoan doi co phieu", "hoan doi trai phieu", "chuyen doi trai phieu", "chuyen doi co phieu", "chuyen quyen"]
  * Trả về "Chứng quyền": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa đồng thời keywords "chung quyen" và một trong các keywords ["dang ky", "giay chung nhan", "khai bao", "dieu chinh", "thuc hien", "do dao han", "huy"]
  * Trả về null: nếu không thoả bất kỳ trường hợp nào vừa nêu.

* Trường [LoaiQuyen]:
  * Nếu [NhomQuyen]="Cổ tức cổ phiếu / Cổ phiếu thưởng":
    * Trả về "Cổ phiếu thưởng": nếu giá trị tại cột title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["phat hanh co phieu", "nhan co phieu", "nhan them co phieu", "co phieu thuong", "thuong co phieu"]; hoặc trong giá trị của một trong 2 trường title hoặc lý_do_mục_đích (được tiền xử lý) xuất hiện đồng thời 2 keywords "co phieu" và "tang von"
    * Trả về "Cổ tức cổ phiếu": nếu giá trị tại trường title hoặc lý_do_mục_đích (được tiền xử lý) chứa một trong số các keywords ["co tuc bang co phieu", "co tuc co phieu", "chuyen doi thanh co phan", "co phieu de tra co tuc"] hoặc đồng thời chứa 2 keywords "co tuc" và "co phieu".
  * Nếu [NhomQuyen]="Đăng ký Lưu ký":
    * Trả về "Đăng ký": Nếu một trong các trường title có chứa keyword "dang ky"
	* Trả về "Lưu ký": Nếu một trong các trường title có chứa keyword "luu ky".
  * Nếu [NhomQuyen]="Tin huỷ", xét nếu một trong các trường title hoặc text_content (được tiền xử lý):
    * có chứa keyword "huy dang ky chung khoan" -> trả ra "Hủy đăng ký chứng khoán"
    * có chứa keyword "huy dang ky chung quyen" -> trả ra "Hủy đăng ký chứng quyền"
    * có chứa keyword "huy dang ky trai phieu" -> trả ra "Hủy đăng ký trái phiếu"
    * có chứa keyword "huy dot chot danh sach thuc hien chung quyen" -> trả ra "Hủy đợt chốt danh sách thực hiện chứng quyền"
    * có chứa keyword "huy danh sach nguoi so huu chung khoan" -> trả ra "Hủy danh sách người sở hữu chứng khoán"
    * có chứa keyword "huy thong bao ngay dang ky cuoi cung" -> trả ra "Hủy thông báo ngày đăng ký cuối cùng".
  * Còn lại, trả về null.

* Trường [MaISIN], lần lượt xét:
  * Nếu giá trị của cột mã_isin không null -> lấy giá trị của cột mã_isin
  * Còn lại -> Trong text_content (được tiền xử lý), tìm keyword "isin" đầu tiên xuất hiện, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra Mã ISIN là một chuỗi có 12 ký tự bắt đầu đầu bằng 1 ký tự chữ và kết thúc bằng 1 ký tự số, sau đó UPPERCASE chuỗi này và trả ra kết quả. Ví dụ, một số Mã ISIN trong thực tế: VN0VHM126016, VNHDB1260174, VNHDB1260166, VN0FUEVFVND5, VN000000PTE4.

* Trường [MaTrongNuoc]: Trong text_content (được tiền xử lý), tìm một trong số các keywords ["ma quyen mua", "ma trong nuoc"] đầu tiên xuất hiện, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra Mã trong nước là một chuỗi có 9 ký tự bắt đầu đầu bằng 1 ký tự chữ và kết thúc bằng 1 ký tự số, sau đó UPPERCASE chuỗi này và trả ra kết quả. Ví dụ, một số Mã trong nước thực tế: MIRSBT261, MIRVTP261, MIRMBS261, MIRGIC261, MIRTDF261.

* Trường [NgayChot]:
  * Nếu giá trị trường ngày_đăng_ký_cuối không null -> Trả ra giá trị của trường ngày_đăng_ký_cuối
  * Nếu giá trị trường ngày_đăng_ký_cuối là null -> Tìm một trong số các keywords ["ngay dang ky cuoi", "thoi gian dang ky cuoi", "ngay chot", "thoi gian chot"] đầu tiên xuất hiện, xét 200 ký tự tiếp theo và lấy ra chuỗi ký tự sớm nhất tìm được có định dạng ngày. Chuỗi ký tự chứa thông tin ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'ngay d thang m nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó.

* Trường [NgayGDKHQ]: Lấy ngày ngay trước [NgayChot] 1 ngày, định dạng dd/mm/yyyy.

* Trường [NgayThucHien] & trường [UocTinhNgayThucHien]:
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là "Quyền biểu quyết" -> Trong text_content (được tiền xử lý), tìm một trong các keywords ["thoi gian thuc hien", "ngay thuc hien", "thoi gian hien", "ngay hien", "thoi gian thuc", "ngay thuc"] xuất hiện đầu tiên, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc "+" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra chuỗi ký tự có định dạng ngày đầu tiên tìm được, tuy nhiên, nếu tìm thấy 2 chuỗi ký tự có định dạng ngày thì lấy chuỗi ở gần cuối hơn. Chuỗi ký tự chứa thông tin ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'ngay d thang m nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó. Nếu không tìm thấy bất kỳ chuỗi có định dạng ngày nào, trả ra giá trị ngày cuối của tháng ngay sau tháng của [NgayChot] 1 tháng (ví dụ, [NgayChot]=22/12/2025, chốt tại tháng 12/2025 thì cần trả ra ngày cuối của tháng ngay sau đó là tháng 01/2026, là 31/01/2026) và đồng thời gán trường [UocTinhNgayThucHien]=1 để đánh dấu. Trường [UocTinhNgayThucHien] để trống trong tất cả các trường hợp còn lại.
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là các giá trị còn lại -> trả về null.

* Trường [NgayThanhToan]:
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là một trong số các loại ["Quyền biểu quyết", "Cổ phiếu thưởng", "Cổ tức bằng cổ phiếu"] -> trả về null
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là các loại còn lại -> Trong text_content (được tiền xử lý), tìm một trong số các keywords ["ngay thuc hien", "ngay thanh toan", "thoi gian thuc hien", "thoi gian thanh toan"], sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra chuỗi ký tự có định dạng ngày đầu tiên tìm được, tuy nhiên, nếu tìm thấy 2 chuỗi ký tự có định dạng ngày thì lấy chuỗi ở gần cuối hơn. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'ngay d thang m nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó.

* Trường [CNQuyenMuaTuNgay] & trường [CNQuyenMuaDenNgay]:
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là "Quyền mua" -> Trong text_content (được tiền xử lý), tìm một trong số các keywords ["thoi gian chuyen nhuong", "ngay chuyen nhuong", "han chuyen nhuong"] xuất hiện đầu tiên, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc "+" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước), lấy chuỗi ký tự có định dạng ngày đầu tiên tìm thấy gán cho trường [CNQuyenMuaTuNgay] và chuỗi ký tự có định dạng ngày sau cùng nhất gán cho trường [CNQuyenMuaDenNgay]. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'ngay d thang m nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là các loại còn lại -> trả về null.
  
* Trường [DKQuyenMuaTuNgay] & trường [DKQuyenMuaDenNgay]:
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là "Quyền mua" -> Trong text_content (được tiền xử lý), tìm một trong số các keywords ["thoi gian dang ky", "ngay dang ky", "han dang ky", "thoi gian dat", "ngay dat", "han dat", "thoi gian nop tien", "ngay nop tien", "han nop tien"] xuất hiện đầu tiên, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc "+" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước), lấy chuỗi ký tự có định dạng ngày đầu tiên tìm thấy gán cho trường [DKQuyenMuaTuNgay] và chuỗi ký tự có định dạng ngày sau cùng nhất gán cho trường [DKQuyenMuaDenNgay]. Chuỗi ký tự có định dạng ngày này có thể là một trong các dạng sau ['dd/mm/YYYY', 'd/m/YYYY', 'd/mm/YYYY', 'dd/m/YYYY', 'mm/YYYY', 'm/YYYY', 'ngay dd thang mm nam YYYY', 'ngay dd thang m nam YYYY', 'ngay d thang mm nam YYYY', 'ngay d thang m nam YYYY', 'thang mm nam YYYY', 'thang m nam YYYY']. Kết quả trả ra là ngày ở định dạng dd/mm/YYYY, nếu chuỗi tìm được rơi vào một trong các dạng ['mm/YYYY', 'm/YYYY', 'thang mm nam YYYY', 'thang m nam YYYY'] thì lấy ngày cuối của tháng đó
  * Nếu giá trị tại trường [NhomQuyen] (như rule bên trên) là các loại còn lại -> trả về null.

* Trường [DonViHuongQuyen] & trường [GiaTriHuongQuyen]:
  * Nếu [NhomQuyen]="Quyền biểu quyết" -> cả 2 trường [DonViHuongQuyen] và [GiaTriHuongQuyen] đều được gán giá trị mặc định là 1
  * Đơn vị hưởng quyền (DonViHuongQuyen) là số cổ phiếu nắm giữ đơn vị (mẫu số của tỷ lệ thực hiện) dùng để xác định một giá trị hưởng quyền tương ứng được quy định (GiaTriHuongQuyen, tử số của tỷ lệ thực hiện). Ví dụ, với "Tỷ lệ thực hiện: 50%/cổ phiếu (1 cổ phiếu được nhận 5.000 đồng)" thì Đơn vị hưởng quyền là 1 và Giá trị hưởng quyền là 5000
  * Phạm vi tìm kiếm 2 thông tin [DonViHuongQuyen] & [GiaTriHuongQuyen] là trong text_content (được tiền xử lý) từ sau một trong các keywords ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"] đầu tiên xuất hiện cho đến khi gặp ký tự "-" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước)
  * Giả sử Đơn vị hưởng quyền là X & Giá trị hưởng quyền là Y thì chúng có thể được tìm thấy từ 1 trong số những mẫu nội dung sau: "<X> co phieu duoc nhan <Y>", "<X> co phieu nhan duoc <Y>", "<X> co phieu duoc <Y>", "<X> co phieu nhan <Y>", "<X> trai phieu duoc nhan <Y>", "<X> trai phieu nhan duoc <Y>", "<X> trai phieu duoc <Y>", "<X> trai phieu nhan <Y>", "<X> co phieu se duoc nhan <Y>", "<X> co phieu se nhan duoc <Y>", "<X> co phieu se duoc <Y>", "<X> co phieu se nhan <Y>", "<X> trai phieu se duoc nhan <Y>", "<X> trai phieu se nhan duoc <Y>", "<X> trai phieu se duoc <Y>", "<X> trai phieu se nhan <Y>", "<X> co phieu se duoc nhan them <Y>", "<X> co phieu se nhan duoc them <Y>", "<X> co phieu se duoc them <Y>", "<X> co phieu se nhan them <Y>", "<X> trai phieu se duoc nhan them <Y>", "<X> trai phieu se nhan duoc them <Y>", "<X> trai phieu se duoc them <Y>", "<X> trai phieu se nhan them <Y>", "<X> co phieu duoc nhan them <Y>", "<X> co phieu nhan duoc them <Y>", "<X> co phieu duoc them <Y>", "<X> co phieu nhan them <Y>", "<X> trai phieu duoc nhan them <Y>", "<X> trai phieu nhan duoc them <Y>", "<X> trai phieu duoc them <Y>", "<X> trai phieu nhan them <Y>", "<X> co phieu - <Y> quyen bieu quyet", "<X> trai phieu - <Y> quyen bieu quyet", "<X> co phieu – <Y> quyen bieu quyet", "<X> trai phieu – <Y> quyen bieu quyet", "<X> co phieu — <Y> quyen bieu quyet", "<X> trai phieu — <Y> quyen bieu quyet", "<X> chung chi quy - <Y> quyen bieu quyet", "<X> chung chi quy – <Y> quyen bieu quyet", "<X> chung chi quy — <Y> quyen bieu quyet", "<X>:<Y>", "<X> : <Y>", "<X>/<Y>", "<X> / <Y>". Ngoài ra, trong những mẫu nội dung vừa nêu, có thể có những trường hợp có thêm mô tả bằng chữ ngay sau X theo dạng "<X> (<mô tả bằng chữ của X>)<...>" hoặc "<X>(<mô tả bằng chữ của X>)<...>"thì cần bắt được X chuẩn xác; ví dụ, trường hợp "01 (một) trái phiếu nhận được 1.051.000.000" hoặc "01(một) trái phiếu nhận được 1.051.000.000" thì cần xác định được X=1. Các số X và Y trong text_content có format với dấu chấm phần ngìn và dấu phẩy thập phân, ví dụ như 4.658.904,110 hay 12.876
  * Đặc biệt lưu ý, nếu Giá trị hưởng quyền sau khi trích xuất là một số thập phân với dấu "," thì cần có thêm 1 bước xử lý để ra được kết quả cuối, đó là nhân Đơn vị hưởng quyền và Giá trị hưởng quyền với cùng một số luỹ thừa của 10 để Giá trị hưởng quyền là một số vừa nguyên và tỷ lệ thực hiện (Giá trị hưởng quyền/ Đơn vị hưởng quyền) không thay đổi so với số gốc. Ví dụ, với trường hợp nội dung "100 cổ phiếu được nhận 73.500,7476" thì thông tin gốc lấy được là Đơn vị hưởng quyền là 100 & Giá trị hưởng quyền là 73500,7476, tuy nhiên, 73500,7476 là một số thập phân với 4 chữ số sau dấu phẩy nên ta cần nhân cả Đơn vị hưởng quyền & Giá trị hưởng quyền cho 1e4 để ra giá trị cuối cùng của Đơn vị hưởng quyền là 1000000 & Giá trị hưởng quyền là 735007476.

* Trường [TyLeMenhGia], là kiểu số thực:
  * Nếu [NhomQuyen] khác "Cổ tức tiền" -> trả về null
  * Nếu [NhomQuyen]="Cổ tức tiền" -> Trong text_content (được tiền xử lý), tìm một trong số các keywords ["ty le thuc hien", "ty le thanh toan", "ti le thuc hien", "ti le thanh toan"] xuất hiện đầu tiên, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc "+" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra con số % xuất hiện đầu tiên (ví dụ, 5% thì trả ra 5, 2.5% thì trả ra 2.5). Trường hợp không tìm thấy con số %, lấy [GiaTriHuongQuyen] chia cho 1e(2+k) với k là số luỹ thừa của 10 để biểu diễn cho [DonViHuongQuyen] về 1 (ví dụ, DonViHuongQuyen là 1 thì k là 0 và trả ra [GiaTriHuongQuyen]/1e2, 100 thì k là 2 và trả ra [GiaTriHuongQuyen]/1e4, 1000 thì k là 3 và trả ra [GiaTriHuongQuyen]/1e5).

* Trường [GiaPhatHanh]: Trong text_content (được tiền xử lý), tìm keyword "gia phat hanh" đầu tiên xuất hiện, sau đó duyệt tiếp đến khi gặp ký tự "-" hoặc "+" hoặc ký tự xuống dòng (tuỳ điều kiện nào xảy ra trước) và lấy ra chuỗi số đầu tiên tìm thấy (chúng thường ở dạng số có dấu chấm phần nghìn như 10.000, 25.000, 100.000, và sẽ cần trả ra các giá trị tương ứng là 10000, 25000, 100000).

* Trường [NoiDung]:
  * Nếu trường [NhomQuyen] (như rule bên trên) là một trong các loại ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Thay đổi"] -> trả ra null
  * Nếu trường [NhomQuyen] (như rule bên trên) là "Tin huỷ", xét nếu một trong các trường title hoặc text_content (được tiền xử lý):
    * có chứa keyword "huy dang ky chung khoan" -> trả ra "Hủy đăng ký chứng khoán"
    * có chứa keyword "huy dang ky chung quyen" -> trả ra "Hủy đăng ký chứng quyền"
    * có chứa keyword "huy dang ky trai phieu" -> trả ra "Hủy đăng ký trái phiếu"
    * có chứa keyword "huy dot chot danh sach thuc hien chung quyen" -> trả ra "Hủy đợt chốt danh sách thực hiện chứng quyền"
    * có chứa keyword "huy danh sach nguoi so huu chung khoan" -> trả ra "Hủy danh sách người sở hữu chứng khoán"
    * có chứa keyword "huy thong bao ngay dang ky cuoi cung" -> trả ra "Hủy thông báo ngày đăng ký cuối cùng"
  * Nếu trường [NhomQuyen] (như rule bên trên) là "Quyền biểu quyết" -> Trả ra giá trị của trường [TieuDe] (như mô tả bên trên)
  * Nếu trường [NhomQuyen] (như rule bên trên) là các loại còn lại -> Lấy giá trị của trường [TieuDe], sau đó, nối thêm vào nó 3 thông tin được lấy trong text_content ở định dạng gốc có dấu Tiếng Việt (trước khi tiền xử lý) và nối với nhau bởi dấu " - " gồm: một là Tỷ lệ thực hiện (hoặc Tỷ lệ thanh toán), hai là Ngày thanh toán (hoặc Ngày thực hiện, Thời gian thực hiện, Thời gian thanh toán), và ba là Giá phát hành.  Ví dụ: code THS có [NhomQuyen]="Cổ tức tiền" thì Giá trị của cột [NoiDung] sẽ là "Chi trả cổ tức bằng tiền năm 2025 - Tỷ lệ thực hiện: 8%/ cổ phiếu (1 cổ phiếu được nhận 800 đồng) - Ngày thanh toán: 15/05/2026 - Giá phát hành: 10.000 đồng". Trong text_content (được tiền xử lý), thông tin Tỷ lệ thực hiện được bắt đầu bằng keyword "ty le thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Tỷ lệ thanh toán được bắt đầu bằng "ty le thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng); Ngày thực hiện được bắt đầu bằng "ngay thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Ngày thanh toán được bắt đầu bằng "ngay thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Thời gian thực hiện được bắt đầu bằng "thoi gian thuc hien" và lấy đến hết dòng (khi gặp ký tự xuống dòng), hoặc Thời gian thanh toán được bắt đầu bằng "thoi gian thanh toan" và lấy đến hết dòng (khi gặp ký tự xuống dòng); Giá phát hành được bắt đầu bằng "gia phat hanh" và lấy đến hết dòng (khi gặp ký tự xuống dòng).

* Trường [is_completed] là một trường boolean trả ra giá trị 0 hoặc 1:
  * Trả ra giá trị 1 khi thoả mãn cả 4 tiêu chí sau:
    * Các trường [NhomQuyen], [MaChungKhoan], [NgayChot] đều không null
    * Nếu [NhomQuyen] là "Quyền mua" -> các trường sau không null: [CNQuyenMuaTuNgay], [CNQuyenMuaDenNgay], [DKQuyenMuaTuNgay], và [DKQuyenMuaDenNgay], [DonViHuongQuyen], [GiaTriHuongQuyen], [GiaPhatHanh], [MaISIN], [MaTrongNuoc]
    * Nếu [NhomQuyen] là "Cổ tức tiền" -> trường [TyLeMenhGia] không null
	* Nếu [NhomQuyen] là "Cổ tức cổ phiếu / Cổ phiếu thưởng" -> các trường sau không null: [DonViHuongQuyen], [GiaTriHuongQuyen]
    * Nếu [NhomQuyen] là "Quyền biểu quyết" -> trường [NgayThucHien] không null
  * Còn lại, trả ra giá trị 0.

* Trường [is_special] là một trường boolean trả ra giá trị 0 hoặc 1:
  * Trả ra 1 khi [NhomQuyen] là một trong các loại ["Hoán đổi chuyển đổi", "Khai báo chứng quyền", "Đăng ký Lưu ký", "Tin huỷ", "Thay đổi"]
  * Còn lại, trả ra giá trị 0.

