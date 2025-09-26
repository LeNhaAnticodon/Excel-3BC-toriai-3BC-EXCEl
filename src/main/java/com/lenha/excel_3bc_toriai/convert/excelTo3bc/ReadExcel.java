package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import javafx.collections.ObservableList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.Map;

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

    public static void readExcel(String fileExcelPath, ObservableList<Map<String[], Map<Double, Integer>>> toriaiSheets) throws FileNotFoundException {
        // lấy địa chỉ file excel
        excelPath = fileExcelPath;

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
            for (Sheet sheet: workbook) {
                // lấy mã vật liệu
                String kousyu = sheet.getRow(HANG_MA_VAT_LIEU).getCell(COT_MA_VAT_LIEU).getStringCellValue();
                // lấy khối lượng riêng
                double khoiLuongRieng;
                Cell khoiLuongRiengCell = sheet.getRow(HANG_KHOI_LUONG_RIENG).getCell(COT_KHOI_LUONG_RIENG);
                khoiLuongRieng = Double.parseDouble(getStringCellValue(khoiLuongRiengCell));

                // tạo mảng chứa khối lượng riêng và mã vật liệu
                String[] kousyuVaKhoiLuongRiengArr = new String[CAC_THONG_SO_KHOI_LUONG_RIENG_VA_MA_VAT_LIEU];

                convertKousyuVaKhoiLuongRiengExcelTo3bc(kousyu, khoiLuongRieng, kousyuVaKhoiLuongRiengArr);

                // lấy hàng cuối cùng chứa dữ liệu trong cột a, chính là hàng cuối cùng chứa chiều dài số lượng sản phẩm
                int lastRowSeihin = getLastRowWithDataInColumn(sheet, COT_CHIEU_DAI_SAN_PHAM); // Cột A = index 0
//                System.out.println(kousyu + " " + lastRowSeihin);


                Map<Double, Integer> seihins = new LinkedHashMap<>();

                // duyệt qua các hàng chứa sản phẩm trong sheet đang duyệt và thêm nó vào map sản phẩm
                for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
                    Row row = sheet.getRow(i);
                    // lấy chiều dài sản phẩm
                    Double seihinZenchou = Double.valueOf(getStringCellValue(row.getCell(0)));
                    // lấy số lượng sản phẩm
                    Integer seihinHonsuu = (int) Double.parseDouble(getStringCellValue(row.getCell(1)));// do có thể kết quả trả về là số thực nên cần chuyển String sang số thực trước rồi mới chuyển sang int

//                    System.out.println(seihinZenchou + " : " + seihinHonsuu);
                    // thêm các thông số sản phẩm vào map
                    seihins.put(seihinZenchou, seihinHonsuu);

                }

                // tạo map chứa toriai và add vật liệu, khối lượng riêng + danh sách các sản phẩm(chiều dài, số lượng) vừa lấy ở trên vào
                LinkedHashMap<String[], Map<Double, Integer>> toriai = new LinkedHashMap<>();
                toriai.put(kousyuVaKhoiLuongRiengArr, seihins);


                // thêm toriai của sheet đang duyệt vào map các toriai
                toriaiSheets.add(toriai);

            }



        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }
    }

    /**
     * chuyển đổi khối lượng riêng và mã vật liệu của Excel sang mảng để ghi vào 3bc
     * @param kousyu mã vật liệu kiểu Excel
     * @param khoiLuongRieng khối lượng riêng của Excel
     * @param kousyuVaKhoiLuongRiengArr mảng để ghi vào 3bc
     */
    private static void convertKousyuVaKhoiLuongRiengExcelTo3bc(String kousyu, double khoiLuongRieng, String[] kousyuVaKhoiLuongRiengArr) {
        String kiHieu3bc = "";

        String[] kousyuMarkArr;
        String[] kousyuSizeArr;

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

        for (String[] capKyHieu: cacCapKyHieuVatLieuExCelVa3bc) {
            String kiHieuExcel = capKyHieu[0];
            if (kousyu.contains(kiHieuExcel)){
                kiHieu3bc = capKyHieu[1];
                themKosyuVaKhoiLuongRiengVaoMangCua3bc(kousyu, khoiLuongRieng, kousyuVaKhoiLuongRiengArr,kiHieuExcel, kiHieu3bc);
            }
        }



    }

    private static void themKosyuVaKhoiLuongRiengVaoMangCua3bc(String kousyu, double khoiLuongRieng, String[] kousyuVaKhoiLuongRiengArr,String kiHieuExcel, String kiHieu3bc) {
        String[] kousyuMarkArr;
        String[] kousyuSizeArr;
        String size1 = null;
        String size2 = null;
        String size3 = null;
        String size4 = "0";

        kousyuMarkArr = kousyu.split(kiHieuExcel);
        kousyuSizeArr = kousyuMarkArr[kousyuMarkArr.length - 1].split("X");

        try{
            size1 = kousyuSizeArr[0];
            size2 = kousyuSizeArr[1];
            size3 = kousyuSizeArr[2];
            size4 = kousyuSizeArr[3];
        }catch (ArrayIndexOutOfBoundsException e){
            System.out.println("phần tử này đã vượt giới hạn chứa các size");
        }

        kousyuVaKhoiLuongRiengArr[0] = String.valueOf(khoiLuongRieng);
        kousyuVaKhoiLuongRiengArr[1] = kiHieu3bc;
        kousyuVaKhoiLuongRiengArr[2] = size1;
        kousyuVaKhoiLuongRiengArr[3] = size2;
        kousyuVaKhoiLuongRiengArr[4] = size3;
        kousyuVaKhoiLuongRiengArr[5] = size4;
        System.out.println(Arrays.toString(kousyuVaKhoiLuongRiengArr));

    }

    private static String getStringCellValue(Cell cell) {
        if (cell != null) {
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
     * Hàm tìm hàng cuối cùng có dữ liệu trong một cột
     * @param sheet Sheet cần kiểm tra
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
