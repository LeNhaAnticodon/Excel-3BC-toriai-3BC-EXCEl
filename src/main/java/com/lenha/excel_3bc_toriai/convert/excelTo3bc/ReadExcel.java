package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import com.lenha.excel_3bc_toriai.dao.SetupData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

public class ReadExcel {
    // link của file excel
    private static String excelPath = "";

    private static final int HANG_MA_VAT_LIEU = 2;
    private static final int COT_MA_VAT_LIEU = 8;
    private static final int HANG_KHOI_LUONG_RIENG = 6;
    private static final int COT_KHOI_LUONG_RIENG = 1;
    private static final int COT_CHIEU_DAI_SAN_PHAM = 0;
    // lấy hàng đầu tiên chứa chiều dài số lượng sản phẩm
    private static final int HANG_DAU_TIEN_CHUA_SAN_PHAM = 9;
    private static final int CAC_THONG_SO_KHOI_LUONG_RIENG_VA_MA_VAT_LIEU = 6;

    public static boolean readExcel(String fileExcelPath, Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets) throws FileNotFoundException {
        // lấy địa chỉ file excel
        excelPath = fileExcelPath;

        // biến kiểm tra trong danh sách các sheet tính vật liệu có sheet nào đó có vật liệu không giống với các vật liệu đã cài đặt sẵn trong chương trình không,
        // nếu có thì sẽ thay vật liệu của toàn bộ các sheet bằng bộ vật liệu tự cho trong danh sách dự phòng đã tạo khi khởi tạo chương trình
        // bộ vật liệu dự phòng lấy từ file excel VAT_LIEU_DU_PHONG.xlsx
        boolean co1VatLieuKhongTonTai = false;

        try (FileInputStream excelFileFis = new FileInputStream(excelPath)) {
            Workbook workbook = new XSSFWorkbook(excelFileFis);
            /*
            test
            // lấy sheet đầu tiên
//            var sheet = workbook.getSheetAt(0);
            Sheet sheet0 = workbook.getSheetAt(0);
            // lấy số lượng hàng
            int rowCount = sheet0.getPhysicalNumberOfRows();
            // lấy số lượng cột
            int colCount = sheet0.getRow(0).getPhysicalNumberOfCells();*/

            // lấy số lượng sheets
            int sheetCount = workbook.getNumberOfSheets();


            // lặp qua các sheet trong excel
            for (Sheet sheet : workbook) {
                // lấy mã vật liệu
                String kousyu = getFullStringCellValue(sheet.getRow(HANG_MA_VAT_LIEU).getCell(COT_MA_VAT_LIEU));
                // lấy khối lượng riêng
                double khoiLuongRieng;
                Cell khoiLuongRiengCell = sheet.getRow(HANG_KHOI_LUONG_RIENG).getCell(COT_KHOI_LUONG_RIENG);
                khoiLuongRieng = Double.parseDouble(getStringNumberCellValue(khoiLuongRiengCell));

                // tạo mảng chứa khối lượng riêng và mã vật liệu
                String[] kousyuVaKhoiLuongRiengArr = new String[CAC_THONG_SO_KHOI_LUONG_RIENG_VA_MA_VAT_LIEU];
                // khởi tạo cho mảng để tránh bị null trong trường hợp vật liệu không có trong bộ vật liệu cho trước
                // thì mảng chứa vật liệu này sẽ không được gán nên cần khởi tạo cho vật liệu các giá trị rỗng
                kousyuVaKhoiLuongRiengArr[0] = "";
                kousyuVaKhoiLuongRiengArr[1] = "";
                kousyuVaKhoiLuongRiengArr[2] = "";
                kousyuVaKhoiLuongRiengArr[3] = "";
                kousyuVaKhoiLuongRiengArr[4] = "";
                kousyuVaKhoiLuongRiengArr[5] = "";
//                System.out.println("khoi tao" + kousyuVaKhoiLuongRiengArr[0]);


                // gọi hàm phân tách các thông số của vật liệu rồi gán nó + khối lượng riêng vào mảng kousyuVaKhoiLuongRiengArr
                // chứa các thông số vật liệu để phục vụ cho chuyển đổi sang 3bc
                // lấy kết quả vật liệu của sheet tính vật liệu này có trong bộ vật liệu cho trước không
                boolean vatLieuTonTai = convertKousyuVaKhoiLuongRiengExcelTo3bc(kousyu, khoiLuongRieng, kousyuVaKhoiLuongRiengArr);

                // nếu sheet này vật liệu không có trong bộ vật liệu cho trước thì cho biến kiểm tra là true
                if (!vatLieuTonTai) {
                    co1VatLieuKhongTonTai = true;
                }

                // lấy hàng cuối cùng chứa dữ liệu trong cột a, chính là hàng cuối cùng chứa chiều dài số lượng sản phẩm
                int lastRowSeihin = getLastRowWithDataInColumn(sheet, COT_CHIEU_DAI_SAN_PHAM); // Cột A = index 0
//                System.out.println(kousyu + " " + lastRowSeihin);

                // map chứa các cặp chiều dài và số lượng sản phẩm của sheet đang duyệt
                List<Map.Entry<Double, Integer>> seihins = new LinkedList<>();
//                Map<Double, Integer> seihins = new LinkedHashMap<>();

                // duyệt qua các hàng chứa sản phẩm trong sheet đang duyệt và thêm nó vào map sản phẩm
                for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
                    Row row = sheet.getRow(i);
                    // lấy chiều dài sản phẩm
                    Double seihinZenchou = Double.valueOf(getStringNumberCellValue(row.getCell(0)));
                    // lấy số lượng sản phẩm
                    Integer seihinHonsuu = (int) Double.parseDouble(getStringNumberCellValue(row.getCell(1)));// do có thể kết quả trả về là số thực nên cần chuyển String sang số thực trước rồi mới chuyển sang int

//                    System.out.println(seihinZenchou + " : " + seihinHonsuu);
                    // thêm các thông số sản phẩm vào map
//                    seihins.put(seihinZenchou, seihinHonsuu);
                    seihins.add(new AbstractMap.SimpleEntry<>(seihinZenchou, seihinHonsuu));

                }

                // thêm toriai của sheet đang duyệt vào map các toriai
                toriaiSheets.put(kousyuVaKhoiLuongRiengArr, seihins);

            }
            System.out.println("có vật liệu không tồn tại trong bộ vật liệu cho trước: " + co1VatLieuKhongTonTai);

            if (toriaiSheets.size() > 100) {
                throw new VerifyError();
            }

            // nếu có vật liệu dự phòng thì thay thế tất cả vật liệu của các sheet tính vật liệu sang bộ vật liệu dự phông
            // phải thay đổi tất cả các vật liệu vì nếu chỉ thay đổi vật liệu không có kia thì vật liệu dự phòng đã thay đổi có thể trùng với
            // vật liệu đang tồn tại trong 1 tính vật liệu của sheet khác, khi này tính vật liệu sẽ không còn đúng nữa
            if (co1VatLieuKhongTonTai) {
                // tạo biến đếm để gọi đúng thứ tự vật liệu dự phòng trong list dự phòng
                // phải dùng đến AtomicInteger chứ ko phải int vì int không dùng được trong biểu thức lamda ở dưới
                AtomicInteger thuTuVatLieuDuPhong = new AtomicInteger();
                // lấy bộ vật liệu dự phòng
                List<String[]> vatLieuDuPhong = SetupData.getInstance().getVatLieuDuPhong();
                // lặp qua bộ vật liệu gốc và thay thế chúng bằng bộ vật liệu dự phòng
                toriaiSheets.forEach((vatLieu, cacSanPham) -> {
//                        System.out.println(Arrays.toString(vatLieu));
                    // lấy ra vật liệu dự phòng theo đúng thứ tự đang duyệt
                    String[] vatLieuMoi = vatLieuDuPhong.get(thuTuVatLieuDuPhong.get());
                    // thay thế
                    vatLieu[0] = vatLieuMoi[0];
                    vatLieu[1] = vatLieuMoi[1];
                    vatLieu[2] = vatLieuMoi[2];
                    vatLieu[3] = vatLieuMoi[3];
                    vatLieu[4] = vatLieuMoi[4];
                    vatLieu[5] = vatLieuMoi[5];
                    // tăng biến nhớ lên 1
                    thuTuVatLieuDuPhong.set(thuTuVatLieuDuPhong.get() + 1);
                });
            }


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        return co1VatLieuKhongTonTai;
    }

    /**
     * chuyển đổi khối lượng riêng và mã vật liệu của Excel sang mảng để ghi vào 3bc
     *
     * @param kousyu                    mã vật liệu kiểu Excel
     * @param khoiLuongRieng            khối lượng riêng của Excel
     * @param kousyuVaKhoiLuongRiengArr mảng để ghi vào 3bc
     */
    private static boolean convertKousyuVaKhoiLuongRiengExcelTo3bc(String kousyu, double khoiLuongRieng, String[] kousyuVaKhoiLuongRiengArr) {
        // kí hiệu vật liệu khi dùng trên 3bc
        String kiHieu3bc = "";

        // tạo list chứa các cặp ký hiệu  gồm mã vật liệu trên excel và mã tương ứng của nó trên 3bc
        ArrayList<String[]> cacCapKyHieuVatLieuExCelVa3bc = new ArrayList<>();
        String[] maVLExcel1 = {"アングル", "L"};
        String[] maVLExcel2 = {"チャンネル", "U"};
        String[] maVLExcel3 = {"H形鋼", "H"};
        String[] maVLExcel4 = {"I形鋼", "H"};
        String[] maVLExcel5 = {"平鋼", "FB"};
        String[] maVLExcel6 = {"軽量溝形鋼", "CA"};
        String[] maVLExcel7 = {"C形鋼", "C"};
        String[] maVLExcel8 = {"角パイプ", "K"};

        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel1);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel2);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel3);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel4);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel5);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel6);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel7);
        cacCapKyHieuVatLieuExCelVa3bc.add(maVLExcel8);

        // biến kiểm tra xem vật liệu có trong bộ vật liệu cho trước không
        boolean vatLieuTonTai = false;
        // duyệt qua list các cặp ký hiệu để xem vật liệu excel hàm truyền vào khớp với cặp ký hiệu nào thì tìm ra được các size của vật liệu đó và
        // ký hiệu kiểu 3bc tương ứng của nó
        // sau đó gọi hàm phân tách vật liệu excel thành các size riêng biệt rồi thêm các size đó + khối lượng riêng + ký hiệu kiêu 3bc vào mảng thông tin
        for (String[] capKyHieu : cacCapKyHieuVatLieuExCelVa3bc) {
            // kí hiệu kiểu excel
            String kiHieuExcel = capKyHieu[0];
            // nếu vật liệu excel có chứa ký hiệu kiểu excel đang lặp thì lấy ra được ký kiệu kiểu 3bc tương ứng trong mảng đang được lặp
            if (kousyu.contains(kiHieuExcel)) {
                // nếu có tồn tại trong bộ vật liệu cho trước thì cho biến nhớ là true
                vatLieuTonTai = true;
                // lấy ra ký hiệu kiểu 3bc tương ứng của nó
                kiHieu3bc = capKyHieu[1];
                // gọi hàm phân tách vật liệu excel thành các size riêng biệt rồi thêm các size đó + khối lượng riêng + ký hiệu kiêu 3bc vào mảng thông tin
                themKosyuVaKhoiLuongRiengVaoMangCua3bc(kousyu, khoiLuongRieng, kousyuVaKhoiLuongRiengArr, kiHieuExcel, kiHieu3bc);
                break;
            }
        }

        // trả về biến kiểm tra
        return vatLieuTonTai;

    }

    /**
     * phân tách vật liệu excel thành các size riêng biệt rồi thêm các size đó + khối lượng riêng + ký hiệu kiêu 3bc vào mảng thông tin
     *
     * @param kousyu                    vật liệu kiểu excel
     * @param khoiLuongRieng            khối lượng riêng
     * @param kousyuVaKhoiLuongRiengArr mảng sẽ chứa tất các các thông số vật liệu cần thiết cho 3bc
     * @param kiHieuExcel               ký hiệu kiểu excel
     * @param kiHieu3bc                 ký hiệu kiểu 3bc
     */
    private static void themKosyuVaKhoiLuongRiengVaoMangCua3bc(String kousyu, double khoiLuongRieng, String[] kousyuVaKhoiLuongRiengArr, String kiHieuExcel, String kiHieu3bc) {
        // mảng sau khi phân tách vật liệu thành các thành phần sau khi tách vật liệu thành các thành phần thông qua kí hiệu kiểu excel
        String[] kousyuMarkArr;
        // mảng chứa các size của vật liệu
        String[] kousyuSizeArr;
        // tạo 4 size vì chỉ có tối đa 4 size
        String size1 = "0";
        String size2 = "0";
        String size3 = "0";
        String size4 = "0";

        // lấy mảng sau khi phân tách vật liệu
        kousyuMarkArr = kousyu.split(kiHieuExcel);
        // lấy mảng chứa các size vật liệu, chính là phần tử cuối cùng của mảng sau khi phân tách vật liệu
        kousyuSizeArr = kousyuMarkArr[kousyuMarkArr.length - 1].split("X");

        // gán giá trị cho các size vừa lấy được
        // do size 4 không chắc có hay không nên cần phải bắt lỗi
        // nếu không có thì size 4 vẫn được gán giá trị từ trước
        try {
            size1 = extractNumberString(kousyuSizeArr[0]);
            size2 = extractNumberString(kousyuSizeArr[1]);
            size3 = extractNumberString(kousyuSizeArr[2]);
            size4 = extractNumberString(kousyuSizeArr[3]);
        } catch (ArrayIndexOutOfBoundsException e) {
            System.out.println("phần tử này đã vượt giới hạn chứa các size");
        }

        // gán các thông số size + khối lượng riêng + ký hiệu kiêu 3bc đã tìm được vào mảng thông tin
        kousyuVaKhoiLuongRiengArr[0] = String.valueOf(khoiLuongRieng);
        kousyuVaKhoiLuongRiengArr[1] = kiHieu3bc;
        kousyuVaKhoiLuongRiengArr[2] = size1;
        kousyuVaKhoiLuongRiengArr[3] = size2;
        kousyuVaKhoiLuongRiengArr[4] = size3;
        kousyuVaKhoiLuongRiengArr[5] = size4;
//        System.out.println(Arrays.toString(kousyuVaKhoiLuongRiengArr));

    }

    /**
     * Trích xuất và ghép các chữ số trong chuỗi theo thứ tự xuất hiện.
     * - Chuyển full-width digits/dot/comma/sign về ASCII.
     * - Bỏ mọi ký tự không phải chữ số (ngoại trừ 1 dấu thập phân đầu tiên).
     * - Bỏ hoàn toàn các dấu + / - (cả full-width) nếu có.
     * - Trả về chuỗi rỗng nếu không có chữ số hợp lệ.
     */
    public static String extractNumberString(String input) {
        if (input == null) return "";

        String s = normalizeFullWidth(input);

        StringBuilder out = new StringBuilder();
        boolean decimalUsed = false;

        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);

            if (c >= '0' && c <= '9') {
                out.append(c);
            } else if ((c == '.' || c == ',') && !decimalUsed) {
                // Dùng '.' làm dấu thập phân nội bộ; nếu gặp ',' thì đã normalize nhưng vẫn xét chung
                out.append('.');
                decimalUsed = true;
            } else {
                // Bỏ qua mọi ký tự khác, bao gồm dấu '+' hoặc '-' (và full-width tương ứng)
            }
        }

        String res = out.toString();

        // Không chấp nhận kết quả rỗng hoặc chỉ là dấu thập phân
        if (res.isEmpty()) return "";
        if (res.equals(".")) return "";

        // Nếu bắt đầu bằng '.' -> thêm '0' trước
        if (res.charAt(0) == '.') res = "0" + res;

        return res;
    }

    /**
     * Chuyển các ký tự full-width thường gặp ở tiếng Nhật về ASCII tương ứng:
     * - '０'..'９' -> '0'..'9'
     * - '．' -> '.'
     * - '，' -> ','
     * - '－' -> '-' (sẽ bị bỏ sau khi normalize)
     * - '＋' -> '+' (sẽ bị bỏ sau khi normalize)
     */
    private static String normalizeFullWidth(String s) {
        if (s == null) return null;
        StringBuilder sb = new StringBuilder(s.length());
        for (char c : s.toCharArray()) {
            if (c >= '０' && c <= '９') {
                sb.append((char) ('0' + (c - '０')));
            } else if (c == '．') {
                sb.append('.');
            } else if (c == '，') {
                sb.append(',');
            } else if (c == '－') {
                sb.append('-');
            } else if (c == '＋') {
                sb.append('+');
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    /**
     * lấy giá trị của cell và trả về dưới dạng String
     *
     * @param cell cell truyền vào
     * @return giá trị chuỗi
     */
    public static String getFullStringCellValue(Cell cell) {
        if (cell != null) {
            // lấy kiểu của cell rồi gọi hàm lấy giá trị tương ứng theo kiểu đó, chuyển giá trị về String và trả về
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC, FORMULA:
                    return String.valueOf(cell.getNumericCellValue()).trim();
                default:
                    System.out.println("Ô không chứa dữ liệu hợp lệ.");
            }
        }
        return null;
    }

    /**
     * lấy giá trị của cell và trả về dưới dạng String chỉ chứa các chữ số
     *
     * @param cell cell truyền vào
     * @return giá trị chuỗi chỉ chứa các chữ số
     */
    private static String getStringNumberCellValue(Cell cell) {
        if (cell != null) {
            // lấy kiểu của cell rồi gọi hàm lấy giá trị tương ứng theo kiểu đó, chuyển giá trị về String và trả về
            switch (cell.getCellType()) {
                case STRING:
                    // gọi hàm chỉ lấy giá trị số trong chuỗi truyền giá trị chuỗi của cell lấy được và trả về chuỗi chỉ chứa chữ số
                    return extractNumberString(cell.getStringCellValue().trim());
                case NUMERIC, FORMULA:
                    return extractNumberString(String.valueOf(cell.getNumericCellValue()).trim());
                default:
                    System.out.println("Ô không chứa dữ liệu hợp lệ.");
            }
        }
        return null;
    }


    /**
     * Hàm tìm hàng cuối cùng có dữ liệu trong một cột
     *
     * @param sheet       Sheet cần kiểm tra
     * @param columnIndex chỉ số cột (0 = cột A)
     * @return chỉ số hàng (0-based). Nếu không có dữ liệu thì trả về -1
     */
    public static int getLastRowWithDataInColumn(Sheet sheet, int columnIndex) {
        // lấy hàng cuối cùng có chứa dữ liệu của sheet
        int lastRowNum = sheet.getLastRowNum();

        for (int i = lastRowNum; i >= 0; i--) {
            // lấy hàng đang duyệt
            Row row = sheet.getRow(i);
            if (row != null) {
                // lấy cell ở cột chỉ định
                Cell cell = row.getCell(columnIndex);
                // đảm bảo ô không rỗng thì trả về index hàng đang duyệt
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    // Kiểm tra nếu ô có dữ liệu (Text, Number, Date...)
                    if (!cell.toString().trim().isEmpty()) {
                        return i;
                    }
                }
            }
        }
        return -1; // Không tìm thấy dữ liệu
    }
}
