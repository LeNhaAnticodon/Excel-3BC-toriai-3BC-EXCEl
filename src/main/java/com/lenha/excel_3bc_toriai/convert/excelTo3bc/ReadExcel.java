package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import javafx.collections.ObservableList;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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
    private static final int HANG_DAU_TIEN_CHUA_SAN_PHAM = 9;

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

                String[] kousyuAndKhoiLuongRieng

                kousyu = convertKousyuAndKhoiLuongRiengExcelTo3bc(kousyu, khoiLuongRieng);

                // lấy hàng cuối cùng chứa dữ liệu trong cột a, chính là hàng cuối cùng chứa chiều dài số lượng sản phẩm
                int lastRowSeihin = getLastRowWithDataInColumn(sheet, COT_CHIEU_DAI_SAN_PHAM); // Cột A = index 0
//                System.out.println(kousyu + " " + lastRowSeihin);

                // lấy hàng đầu tiên chứa chiều dài số lượng sản phẩm
                int beginRowSeihin = HANG_DAU_TIEN_CHUA_SAN_PHAM;
                Map<Double, Integer> seihins = new LinkedHashMap<>();

                // duyệt qua các hàng chứa sản phẩm trong sheet đang duyệt và thêm nó vào map sản phẩm
                for (int i = beginRowSeihin; i <= lastRowSeihin; i++) {
                    Row row = sheet.getRow(i);
                    // lấy chiều dài sản phẩm
                    Double seihinZenchou = Double.valueOf(getStringCellValue(row.getCell(0)));
                    // lấy số lượng sản phẩm
                    Integer seihinHonsuu = Integer.valueOf(getStringCellValue(row.getCell(1)));

//                    System.out.println(seihinZenchou + " : " + seihinHonsuu);
                    // thêm các thông số sản phẩm vào map
                    seihins.put(seihinZenchou, seihinHonsuu);

                }

                // tạo map chứa toriai và add vật liệu + danh sách các sản phẩm(chiều dài, số lượng) vừa lấy ở trên vào
                LinkedHashMap<String, Map<Double, Integer>> toriai = new LinkedHashMap<>();
                toriai.put(kousyu, seihins);

                // thêm toriai của sheet đang duyệt vào map các toriai
                toriaiSheets.add(new LinkedHashMap<>(toriai));

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

    private static String getStringCellValue(Cell cell) {
        if (cell != null) {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue().trim();
                case NUMERIC, FORMULA:
                    return String.valueOf(cell.getNumericCellValue());
                default:
                    System.out.println("Ô không chứa dữ liệu hợp lệ.");
            }
        }
        return null;
    }


    /**
     * chuyển đổi mã vật liệu kiểu Excel sang kiểu của 3bc
      * @param kousyu mã vật liệu kiểu Excel
     * @return mã vật liệu kiểu 3bc
     */
    private static String convertKousyuAndKhoiLuongRiengExcelTo3bc(String kousyu) {

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
