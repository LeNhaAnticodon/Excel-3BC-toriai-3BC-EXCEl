package com.lenha.excel_3bc_toriai.convert;

import com.lenha.excel_3bc_toriai.convert.copyRowSheetToSheet.RowRangeCopier;
import com.lenha.excel_3bc_toriai.model.ExcelFile;
import com.opencsv.CSVWriter;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeoutException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static com.lenha.excel_3bc_toriai.convert.excelTo3bc.ReadExcel.*;

public class ReadPDFToExcel {

    // list các file(map chứa tính vật liệu) của vật liệu hiện tại
    private static final List<Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>>> fileList = new ArrayList<>();
    // time tháng và ngày
    private static String shortNouKi = "";
    // 備考
    private static String kouJiMe = "";
    // 客先名
    private static String kyakuSakiMei = "";
    // 3 kích thước của vật liệu
    private static int size1;
    private static int size2;
    private static int size3 = 0;
    // ký hiệu loại vật lệu
    private static String koSyuNumMark = "3";
    // 切りロス
    private static String kirirosu = "";

    public static boolean kirirosu20 = true;

    // tên file chl sẽ tạo được ghi trong phần 工事名, chưa bao gồm loại vật liệu
    private static String fileChlName = "";

    // link của file pdf
    private static String pdfPath = "";

    // link thư mục của file excel xlsx sẽ tạo
    private static String xlsxExcelPath = "";
    // link thư mục của file excel csv sẽ tạo
    private static String excelDirPath = "";
    // link thư mục của file chl sẽ tạo
    private static String chlDirPath = "";
    // đếm số dòng sẽ tạo trên file chl
    private static int rowToriAiNum;

    // loại vật liệu và kích thước
    private static String kouSyu;

    // tên loại vật liệu
    private static String kouSyuName;
    // tên file chl đầy đủ sẽ tạo đã bao gồm tên loại vật liệu
    public static String fileName;

    // tổng chiều dài các kozai
    private static double kouzaiChouGoukei = 0;
    private static double seiHinChouGoukei = 0;

    private static String excelCopyPath;

    // tên file excel sẽ tạo được ghi trong phần 工事名, chưa bao gồm loại vật liệu
    public static String fileExcelName = "name";
    private static String bikou = "";
    private static String chuyuBan = "";
    private static String teiHaiSha = "";

    private static int soBozaiMax = 0;
    private static int maxLastColNum = 0;
    private static final Map<Double, Integer> seiHinMap = new LinkedHashMap<>();
    // list chứa danh sách các sản phẩm không trùng lặp
    private static ObservableList<Double> seiHinList = FXCollections.observableArrayList();
    public static File copyFile;

    /**
     * chuyển đổi pdf tính vật liệu thành các file chl theo từng vật liệu khác nhau
     *
     * @param fileExcelRootPath link file excel gốc
     * @param filePDFPath       link file pdf
     * @param fileExcelDirPath  link thư mục chứa file chl sẽ tạo
     * @param excelFileNames    list chứa danh sách các file chl đã tạo
     */
    public static void convertPDFToExcel(String fileExcelRootPath, String filePDFPath, String fileExcelDirPath, ObservableList<ExcelFile> excelFileNames) throws FileNotFoundException, TimeoutException, IOException {
/*        excelFileNames.add(new ExcelFile("test.", "", 0, 0));
        fileName = "test.sysc2";
        throw new TimeoutException();*/

        // xóa danh sách cũ trước khi thực hiện, tránh bị ghi chồng lên nhau
        excelFileNames.clear();

        // lấy địa chỉ file pdf
        pdfPath = filePDFPath;
        // lấy đi chỉ thư mục chứa file excel
//        csvExcelDirPath = fileCSVDirPath;
        // lấy đi chỉ thư mục chứa file excel
        excelDirPath = fileExcelDirPath;
        // lấy đi chỉ thư mục chứa chl
        chlDirPath = fileExcelDirPath;

        // lấy mảng chứa các trang
        String[] kakuKouSyu = getFullToriaiText();
        // lấy trang đầu tiên và lấy ra các thông tin của đơn như tên khách hàng, ngày tháng
        getHeaderData(kakuKouSyu[0]);

        // chuyển mảng các trang sang dạng list
        List<String> kakuKouSyuList = new LinkedList<>(Arrays.asList(kakuKouSyu));

        // kích thước list
        int kakuKouSyuListSize = kakuKouSyuList.size();
        // lặp qua các trang gộp các trang cùng loại vật liệu làm 1 và xóa các trang đã được gộp vào trang khác đi
        for (int i = 1; i < kakuKouSyuListSize; i++) {
            // lấy tên vật liệu đang lặp
            String KouSyuName = extractValue(kakuKouSyuList.get(i), "法:", "梱包");

            // duyệt các trang phía sau, nếu vật liệu giống trang đang lặp thì gộp trang đó vào trang này
            // và xóa trang đó đi
            for (int j = i + 1; j < kakuKouSyuListSize; j++) {
                String KouSyuNameAfter = extractValue(kakuKouSyuList.get(j), "法:", "梱包");
                if (KouSyuName.equals(KouSyuNameAfter)) {
                    kakuKouSyuList.set(i, kakuKouSyuList.get(i).concat(kakuKouSyuList.get(j)));
                    kakuKouSyuList.remove(j);
                    j--;
                    kakuKouSyuListSize--;
                }
            }

            /*if (i > 1) {
                String KouSyuNameBefore = extractValue(kakuKouSyuList.get(i - 1), "法:", "梱包");

                if (KouSyuName.equals(KouSyuNameBefore)) {
                    kakuKouSyuList.set(i - 1, kakuKouSyuList.get(i - 1).concat(kakuKouSyuList.get(i)));
                    kakuKouSyuList.remove(i);
                    i--;
                    kakuKouSyuListSize--;
                }
            }*/
        }
        // BEGIN
        // đoạn code copy file này khác với app chl vì nó chỉ tạo 1 file nên chỉ chạy 1 lần ở đoạn đầu này
        // tạo path chứa file excel
        // mà không chạy trong vòng lặp bên dưới như trong hàm writeDataToChl
        excelCopyPath = excelDirPath + "\\" + fileExcelName + "-NC" + ".xlsx";
        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(excelCopyPath);
        // Kiểm tra nếu file tồn tại thì xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
                System.out.println("File đã được xóa thành công.");
            } else {
                System.out.println("Xóa file thất bại.");
            }
        }
        // path chứa địa chỉ file sẽ được dán từ file copy
        Path copyFilePath = Paths.get(excelCopyPath);

        copyFile = new File(excelCopyPath);

        // Đọc file mẫu từ resources rồi copy file ra địa chỉ của copyFilePath
        try (InputStream sourceFile = ReadPDFToExcel.class.getResourceAsStream("/com/lenha/excel_3bc_toriai/sampleFiles/sample files3.xlsx")) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }
            Files.copy(sourceFile, copyFilePath);
        } catch (IOException e) {
            e.printStackTrace();
            throw new FileNotFoundException();
        }


        // cách cũ clone file từ file gốc nên nó copy cả lỗi của file gốc nên không nên dùng
/*        // Đọc file excel gốc rồi tạo file copy ra địa chỉ theo link copyFilePath
        try (InputStream sourceFile = new FileInputStream(fileExcelRootPath)) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }
            Files.copy(sourceFile, copyFilePath);
        } catch (IOException e) {
            e.printStackTrace();
            throw new FileNotFoundException();
        }*/

        // copy file bằng hàm mới không phải clone file như cách trên để tránh copy cả lỗi của file gốc
//        createNewFileAndCopyValue(new File(fileExcelRootPath), copyFile);
        copyDataToCopyFile(new File(fileExcelRootPath), copyFile);

        // thêm tên file vào list các sheet của file để hiển thị tên file
        excelFileNames.add(new ExcelFile("EXCEL: " + fileExcelName + ".xlsx", "", 0, 0));

/*        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(excelPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
                System.out.println("File is set to read-only.");
            } else {
                System.out.println("Failed to set file to read-only.");
            }
        } else {
            System.out.println("File does not exist.");
        }*/

        // tạo luồng đọc ghi file
        try (FileInputStream fileExcel = new FileInputStream(excelCopyPath)) {
            Workbook workbook = new XSSFWorkbook(fileExcel);

            // lặp qua từng loại vật liệu trong list và ghi chúng vào các file excel
            for (int i = 1; i < kakuKouSyuList.size(); i++) {
                // tách các đoạn bozai thành mảng
                String[] kakuKakou = kakuKouSyuList.get(i).split("加工No:");

                // tại đoạn đầu tiên sẽ không chứa bozai mà chứa tên vật liệu
                // lấy ra thông số loại vật liệu và 3 size riêng lẻ của vật liệu
                getKouSyu(kakuKakou);
                // tạo map kaKouPairs và nhập thông tin tính vật liệu vào
                // kaKouPairs là map chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
                // còn value của kaKouPairs cũng là map chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
                Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = getToriaiData(kakuKakou);

//            writeDataToExcel(kaKouPairs, i - 1, excelFileNames);
//            writeDataToCSV(kaKouPairs, i - 1, excelFileNames);
                // ghi thông tin vào file định dạng .xlsx là file của excel
//            writeDataToChl(kaKouPairs, i, excelFileNames);
                writeDataToExcelToriai2(kaKouPairs, i, excelFileNames, workbook);

                /*workbook.setForceFormulaRecalculation(true);
                try (FileOutputStream fileOut = new FileOutputStream(excelCopyPath)) {
                    workbook.write(fileOut);

                }*/
            }

            // ẩn sheet mẫu
//            workbook.removeSheetAt(0);

            // Yêu cầu Excel tính toán lại tất cả các công thức khi tệp được mở
            workbook.setForceFormulaRecalculation(true);

            // tạo khoảng cách các trang in là 43 dòng
            int numRowPage = 43;
            // số trang đã tạo
            int pageCount = 1;
            // lấy sheet cần dán giá trị các sheet khác vào
            Sheet sheet0 = workbook.getSheetAt(0);
            // lấy số sheet tối đa có trong file excel
            int maxSheet = workbook.getNumberOfSheets();
            // lặp qua các sheet cần copy, copy giá trị, công thức rồi dán vào sheet 0
            for (int i = 1; i < maxSheet; i++) {
                Sheet sheeti = workbook.getSheetAt(i);
                // lấy hàng cuối cùng chứa dữ liệu của sheet copy
                int lastRowi = sheeti.getLastRowNum();
                // lấy hàng cuối cùng chứa dữ liệu của sheet 0
                int lastRow0 = sheet0.getLastRowNum();
                // nếu sheet copy không có hàng nào có dữ liệu thì thoát lặp
                if (lastRowi == 0) {
                    break;
                }
                // nếu số hàng sản phẩm trong sheet gốc mà lớn hơn số hàng đang nhớ theo biến số trang
                // thì tăng biến nhớ số trang lên 1 để số số hàng theo biến nhớ phải lớn hơn số hàng chứ dữ liệu thực tế
                // tăng biến nhớ đến khi nào lớn hơn thực tế
                while (lastRow0 > numRowPage * pageCount) {
                    pageCount++;
                }

                // nếu số hàng sheet gốc + sheet copy lớn hơn số dòng theo biến nhớ trang thì
                // cho hàng dán giá trị sheet copy bắt đầu ở trang mới
                // còn không thì vẫn dán ở hàng cuối chứa dữ liệu của sheet gốc như bình thường
                // lưu ý vị trí hàng + 2 để có khoảng cách cho dễ nhìn
                if (lastRow0 + lastRowi + 2> numRowPage * pageCount) {
                    RowRangeCopier.copyRowRange((XSSFWorkbook) workbook, i, 0, lastRowi, 0, numRowPage * pageCount + 2);
                    pageCount++;
                    continue;
                }

                RowRangeCopier.copyRowRange((XSSFWorkbook) workbook, i, 0, lastRowi, 0, lastRow0 + 2);

            }

            // sau khi đã gộp các sheet thì ẩn các cột không cần hiển thị theo sheet có nhiều bozai nhất
            for (int i = soBozaiMax * 2 + 6; i <= maxLastColNum; i++) {
                sheet0.setColumnHidden(i, true);
            }

            // lấy biến in
            PrintSetup printSetup = sheet0.getPrintSetup();

            // Khổ dọc, giấy A4
            printSetup.setLandscape(false);
            printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);

            // bật fit to page, tự đọng co dãn theo giá trị
            sheet0.setFitToPage(true);

            // tạo khoảng in là mỗi 43 hàng
            int rowsPerPage = 42;
            // lấy hàng cuối chứa dữ liệu
            int lastRow = sheet0.getLastRowNum();
            // tạo các khoảng in mỗi 43 hàng
            for (int r = rowsPerPage; r <= lastRow; r += rowsPerPage) {
                sheet0.setRowBreak(r);
            }

            // ẩn hoặc xóa các sheet đã được gộp vào sheet đầu tiên
            for (int i = 1; i < maxSheet; i++) {
                workbook.setSheetHidden(i, true);
//                workbook.removeSheetAt(1);
            }

            try (FileOutputStream fileOut = new FileOutputStream(excelCopyPath)) {
                workbook.write(fileOut);
                // xóa file gây lỗi khi excel không tự động tính toán lại công thức
                writeWithoutCalcChain(workbook, new File(excelCopyPath));

                workbook.close();
            }

        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // đoạn code cũ dùng cho tạo file chl hoặc csv
        /*// tạo số thứ tự khi ghi tên là thời gian ở ô tên trong file chl để tránh trùng thời gian
        int j = 0;
        // lặp qua từng loại vật liệu trong list và ghi chúng vào các file chl
        for (int i = 1; i < kakuKouSyuList.size(); i++) {
            // tách các đoạn bozai thành mảng
            String[] kakuKakou = kakuKouSyuList.get(i).split("加工No:");

            // tại đoạn đầu tiên sẽ không chứa bozai mà chứa tên vật liệu
            // lấy ra thông số loại vật liệu và 3 size riêng lẻ của vật liệu
            getKouSyu(kakuKakou);
            // tạo list fileList chứa các map và nhập thông tin tính vật liệu vào
            // map chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
            // còn value của kaKouPairs cũng là map chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
            // mỗi map này trong list sẽ tạo thành 1 file trong trường hợp chia file
            List<Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>>> fileList = getToriaiData(kakuKakou);

//            writeDataToExcel(kaKouPairs, i - 1, excelFileNames);
//            writeDataToCSV(kaKouPairs, i - 1, excelFileNames);
            // ghi thông tin của vật liệu này vào các file định dạng sysc2 là file của chl
            int fileListSize = fileList.size();

            //reset lại các tổng chiều dài trước khi ghi các file của vật liệu mới
            kouzaiChouGoukei = 0;
            seiHinChouGoukei = 0;
            for (int k = 0; k < fileListSize; k++) {
                Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = fileList.get(k);
                j++;
                // thêm trong trường hợp số file của vật liệu này lớn hơn 1 thì thêm k vào là hậu tố của file ở hàm writeDataToChl
                writeDataToChl(kaKouPairs, j, excelFileNames, fileListSize, k + 1);
            }
        }*/

    }

    /**
     * chatGPT AI
     * tạo file excel mới và copy y nguyên value từ file excel gốc sang file mới
     * mục đích để sửa lỗi của file gốc, do file gốc tính vật liệu có trường bị lỗi gì đó chưa tìm được nguyên nhân
     * khiến code nhập tính vật liệu không chạy được
     * hàm copyFile của class File nó chỉ clone từ file gốc sang file copy nên file copy vẫn bị lỗi nên không dùng được và phải dùng method này
     *
     * @param srcFile file gốc
     * @param dstFile file copy
     */
    private static void createNewFileAndCopyValue(File srcFile, File dstFile) throws IOException {
        try (FileInputStream fis = new FileInputStream(srcFile);
             XSSFWorkbook srcWb = new XSSFWorkbook(fis);
             XSSFWorkbook dstWb = new XSSFWorkbook()) {

            Map<Integer, XSSFCellStyle> styleMap = new HashMap<>(); // cache cloned styles

            for (int i = 0; i < srcWb.getNumberOfSheets(); i++) {
                XSSFSheet srcSheet = srcWb.getSheetAt(i);
                XSSFSheet dstSheet = dstWb.createSheet(srcSheet.getSheetName());

                // copy column widths: compute max column index conservatively
                int maxCol = 0;
                for (int r = srcSheet.getFirstRowNum(); r <= srcSheet.getLastRowNum(); r++) {
                    Row rr = srcSheet.getRow(r);
                    if (rr != null && rr.getLastCellNum() > maxCol) {
                        maxCol = rr.getLastCellNum();
                    }
                }
                for (int c = 0; c <= maxCol; c++) {
                    try {
                        dstSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
                    } catch (Exception e) {
                        // ignore columns that don't exist in source
                    }
                }

                // copy merged regions
                for (int m = 0; m < srcSheet.getNumMergedRegions(); m++) {
                    CellRangeAddress cra = srcSheet.getMergedRegion(m);
                    dstSheet.addMergedRegion(new CellRangeAddress(cra.getFirstRow(), cra.getLastRow(), cra.getFirstColumn(), cra.getLastColumn()));
                }

                // copy rows and cells
                for (int r = srcSheet.getFirstRowNum(); r <= srcSheet.getLastRowNum(); r++) {
                    XSSFRow srcRow = srcSheet.getRow(r);
                    if (srcRow == null) continue;
                    XSSFRow dstRow = dstSheet.createRow(r);
                    dstRow.setHeight(srcRow.getHeight());

                    short firstCell = srcRow.getFirstCellNum();
                    short lastCell = srcRow.getLastCellNum();
                    if (firstCell < 0 || lastCell < 0) continue;

                    for (int c = firstCell; c < lastCell; c++) {
                        XSSFCell srcCell = srcRow.getCell(c);
                        if (srcCell == null) continue;
                        XSSFCell dstCell = dstRow.createCell(c);

                        // clone style (cache by hashcode)
                        XSSFCellStyle srcStyle = srcCell.getCellStyle();
                        if (srcStyle != null) {
                            int key = srcStyle.hashCode();
                            XSSFCellStyle newStyle = styleMap.get(key);
                            if (newStyle == null) {
                                newStyle = dstWb.createCellStyle();
                                newStyle.cloneStyleFrom(srcStyle);
                                styleMap.put(key, newStyle);
                            }
                            dstCell.setCellStyle(newStyle);
                        }

                        // copy comment (shallow)
                        if (srcCell.getCellComment() != null) {
                            dstCell.setCellComment(srcCell.getCellComment());
                        }

                        // copy value / formula / type
                        switch (srcCell.getCellType()) {
                            case STRING:
                                dstCell.setCellValue(srcCell.getRichStringCellValue());
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(srcCell)) {
                                    dstCell.setCellValue(srcCell.getDateCellValue());
                                } else {
                                    dstCell.setCellValue(srcCell.getNumericCellValue());
                                }
                                break;
                            case BOOLEAN:
                                dstCell.setCellValue(srcCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                try {
                                    dstCell.setCellFormula(srcCell.getCellFormula());
                                } catch (Exception e) {
                                    // If copying formula fails, fallback to cached value (if any)
                                    if (srcCell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                        dstCell.setCellValue(srcCell.getNumericCellValue());
                                    } else if (srcCell.getCachedFormulaResultType() == CellType.STRING) {
                                        dstCell.setCellValue(srcCell.getStringCellValue());
                                    }
                                }
                                break;
                            case BLANK:
                                dstCell.setBlank();
                                break;
                            case ERROR:
                                dstCell.setCellErrorValue(srcCell.getErrorCellValue());
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(dstFile)) {
                dstWb.write(fos);
            }
        }
    }

    public static void copyDataToCopyFile(File srcFile, File destFile) throws IOException {
        try (FileInputStream srcFis = new FileInputStream(srcFile);
             XSSFWorkbook srcWb = new XSSFWorkbook(srcFis)) {

            XSSFWorkbook destWb;
            boolean destExisted = destFile.exists();

            if (destExisted) {
                try (FileInputStream destFis = new FileInputStream(destFile)) {
                    destWb = new XSSFWorkbook(destFis);
                }
                // Ensure only sheet index 0 exists: remove others
                int total = destWb.getNumberOfSheets();
                for (int i = total - 1; i >= 1; i--) { // remove all except 0
                    destWb.removeSheetAt(i);
                }
                // Ensure at least one sheet exists
                if (destWb.getNumberOfSheets() == 0) {
                    destWb.createSheet("Sheet1");
                }
            } else {
                destWb = new XSSFWorkbook();
                destWb.createSheet("Sheet1");
            }

            int srcSheetCount = srcWb.getNumberOfSheets();

            // Create additional sheets in dest by cloning sheet index 0 until counts match
            while (destWb.getNumberOfSheets() < srcSheetCount) {
                destWb.cloneSheet(0); // clone internal sheet 0
            }

            // Now for each sheet index: copy content from src sheet i to dest sheet i
            for (int i = 0; i < srcSheetCount; i++) {
                XSSFSheet srcSheet = srcWb.getSheetAt(i);
                XSSFSheet destSheet = destWb.getSheetAt(i);

                // Optionally rename destination sheet to match source name (ensure unique)
                String desiredName = srcSheet.getSheetName();
                try {
                    destWb.setSheetName(i, uniqueSheetName(destWb, desiredName, i));
                } catch (IllegalArgumentException ex) {
                    // ignore if name invalid or duplicate; we will leave default name
                }

                // Clear destination sheet content (rows and merged regions) but keep column widths if you prefer.
                clearSheetContent(destSheet);

                // Copy sheet content from src to dest (values, formulas, styles, merged regions, column widths)
                copySheetContent(srcSheet, destSheet, destWb);
            }

            // Write back to destFile (overwrite or create)
            try (FileOutputStream fos = new FileOutputStream(destFile)) {
                destWb.write(fos);
            }

            destWb.close();
        }
    }

    // Ensure sheet name is unique; if conflicts, append suffix
    private static String uniqueSheetName(XSSFWorkbook wb, String base, int indexToIgnore) {
        String name = base;
        int suffix = 1;
        while (true) {
            int idx = wb.getSheetIndex(name);
            if (idx == -1 || idx == indexToIgnore) return name;
            name = base + "_" + suffix++;
            if (suffix > 1000) return base + "_" + System.currentTimeMillis();
        }
    }

    // Remove rows and merged regions/content from destination sheet
    private static void clearSheetContent(XSSFSheet sheet) {
        // remove rows
        int first = sheet.getFirstRowNum();
        int last = sheet.getLastRowNum();
        for (int r = last; r >= first; r--) {
            XSSFRow row = sheet.getRow(r);
            if (row != null) sheet.removeRow(row);
        }
        // remove merged regions
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            sheet.removeMergedRegion(i);
        }
    }

    // Copy sheet content (core logic)
    private static void copySheetContent(XSSFSheet srcSheet, XSSFSheet destSheet, XSSFWorkbook destWb) {
        Map<Integer, CellStyle> styleMap = new HashMap<>();

        // copy column widths: determine max column index present in source
        int maxCol = 0;
        for (int r = srcSheet.getFirstRowNum(); r <= srcSheet.getLastRowNum(); r++) {
            Row rr = srcSheet.getRow(r);
            if (rr != null && rr.getLastCellNum() > maxCol) {
                maxCol = rr.getLastCellNum();
            }
        }
        for (int c = 0; c <= maxCol; c++) {
            try {
                destSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
                // also copy hidden state
                destSheet.setColumnHidden(c, srcSheet.isColumnHidden(c));
            } catch (Exception e) {
                // ignore columns that may not exist
            }
        }

        // copy merged regions
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress cra = srcSheet.getMergedRegion(i);
            destSheet.addMergedRegion(new CellRangeAddress(cra.getFirstRow(), cra.getLastRow(), cra.getFirstColumn(), cra.getLastColumn()));
        }

        // copy rows & cells
        for (int r = srcSheet.getFirstRowNum(); r <= srcSheet.getLastRowNum(); r++) {
            XSSFRow srcRow = srcSheet.getRow(r);
            if (srcRow == null) continue;
            XSSFRow destRow = destSheet.createRow(r);
            destRow.setHeight(srcRow.getHeight());

            short firstCell = srcRow.getFirstCellNum();
            short lastCell = srcRow.getLastCellNum();
            if (firstCell < 0 || lastCell < 0) continue;

            for (int c = firstCell; c < lastCell; c++) {
                XSSFCell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                XSSFCell destCell = destRow.createCell(c);

                // copy style (clone into dest workbook)
                XSSFCellStyle srcStyle = srcCell.getCellStyle();
                if (srcStyle != null) {
                    int key = srcStyle.hashCode();
                    XSSFCellStyle newStyle = (XSSFCellStyle) styleMap.get(key);
                    if (newStyle == null) {
                        newStyle = destWb.createCellStyle();
                        try {
                            newStyle.cloneStyleFrom(srcStyle);
                        } catch (Exception e) {
                            // fallback: skip clone if incompatible
                        }
                        styleMap.put(key, newStyle);
                    }
                    destCell.setCellStyle(newStyle);
                }

                // copy comment (shallow)
                if (srcCell.getCellComment() != null) {
                    destCell.setCellComment(srcCell.getCellComment());
                }

                // copy value / type / formula
                switch (srcCell.getCellType()) {
                    case STRING:
                        destCell.setCellValue(srcCell.getRichStringCellValue());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(srcCell)) {
                            destCell.setCellValue(srcCell.getDateCellValue());
                        } else {
                            destCell.setCellValue(srcCell.getNumericCellValue());
                        }
                        break;
                    case BLANK:
                        destCell.setBlank();
                        break;
                    case BOOLEAN:
                        destCell.setCellValue(srcCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        try {
                            destCell.setCellFormula(srcCell.getCellFormula());
                        } catch (Exception e) {
                            // nếu không đặt được formula (ví dụ invalid), fallback to cached result if any
                            try {
                                if (srcCell.getCachedFormulaResultType() == CellType.NUMERIC) {
                                    destCell.setCellValue(srcCell.getNumericCellValue());
                                } else if (srcCell.getCachedFormulaResultType() == CellType.STRING) {
                                    destCell.setCellValue(srcCell.getStringCellValue());
                                }
                            } catch (Exception ex) {
                                destCell.setBlank();
                            }
                        }
                        break;
                    case ERROR:
                        destCell.setCellErrorValue(srcCell.getErrorCellValue());
                        break;
                    default:
                        break;
                }
            }
        }
    }

    /**
     * lấy toàn bộ text của file pdf
     *
     * @return mảng chứa các trang của file pdf, đầu trang chứa tên vật liệu
     */
    private static String[] getFullToriaiText() throws IOException {
        // khởi tạo mảng, có thể ko cần nếu sau đó nó có thể được gán bằng mảng khác
        String[] kakuKouSyu = new String[0];
        // dùng thư viện đọc file pdf lấy toàn bộ text của file
        try (PDDocument document = PDDocument.load(new File(pdfPath))) {
            // nếu file không được mã hóa thì mới lấy được text
            if (!document.isEncrypted()) {
                PDFTextStripper pdfStripper = new PDFTextStripper();
                String toriaiText = pdfStripper.getText(document);

                // chia thành các trang thông qua đoạn 材寸, mỗi trang sẽ chứa loại vật liệu ở đầu trang
                kakuKouSyu = toriaiText.split("材寸");

//                System.out.println(toriaiText);

            }
        }

        return kakuKouSyu;
    }

    /**
     * lấy các thông tin của đơn như ngày, tháng, tên và ghi vào các biến nhớ toàn cục
     * các thông tin nằm trong vùng xác định, dùng hàm extractValue để lấy
     *
     * @param header text chứa thông tin
     */
    private static void getHeaderData(String header) {
        String nouKi = extractValue(header, "期[", "]");
        String[] nouKiArr = nouKi.split("/");
        shortNouKi = nouKiArr[0] + nouKiArr[1] + nouKiArr[2];

        bikou = extractValue(header, "考[", "]");
        kyakuSakiMei = extractValue(header, "客先名[", "]");
        String names = extractValue(header, "工事名[", "]");
        String[] namesArr = names.split("\\+");
        if (namesArr.length == 3) {
            fileExcelName = namesArr[0];
            chuyuBan = namesArr[1];
            teiHaiSha = namesArr[2];
        } else {
            fileExcelName = names;
        }

        System.out.println(shortNouKi + " : " + bikou + " : " + kyakuSakiMei + " : " + chuyuBan + " : " + teiHaiSha);
    }

    /**
     * lấy thông số đầy đủ của vật liệu, tên vật liệu, mã vật liệu, 3 size của vật liệu và ghi vào biến toàn cục
     *
     * @param kakuKakou mảng chứa các tính vật liệu của vật liệu đang xét
     */
    private static void getKouSyu(String[] kakuKakou) {

        // lấy loại vật liệu tại mảng 0 và tách mảng 0 thành các dòng rồi lấu dòng đầu tiên
        // tại dòng này lấy loại vật liệu trong đoạn "法:", "梱包"
        kouSyu = extractValue(kakuKakou[0].split("\n")[0], "法:", "梱包");
        // phân tách vật liệu thành các đoạn thông tin
        String[] kouSyuNameAndSize = kouSyu.split("-");
        // lấy tên vật liệu tại index 0
        kouSyuName = kouSyuNameAndSize[0].trim();

        // từ tên vật liệu lấy ra được  số đại diện cho nó
        switch (kouSyuName) {
            case "K":
                koSyuNumMark = "3";
                break;
            case "L":
                koSyuNumMark = "4";
                break;
            case "FB":
                koSyuNumMark = "5";
                break;
            case "[":
                koSyuNumMark = "6";
                break;
            case "C":
                koSyuNumMark = "7";
                break;
            case "H":
                koSyuNumMark = "8";
                break;
            case "CA":
                koSyuNumMark = "9";
                break;
        }

        // lấy đoạn thông tin 2 chứa các size của vật liệu và phân tách nó thành mảng chứa các size này
        String[] koSyuSizeArr = kouSyuNameAndSize[1].split("x");

        size1 = 0;
        size2 = 0;
        size3 = 0;

        // với từng loại vật liệu có số lượng size khác nhau thì sẽ ghi khác nhau, do chỉ cần thông tin của 3 size và x10
        // size thừa sẽ không cần ghi
        if (koSyuSizeArr.length == 3) {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
            size3 = convertStringToIntAndMul(koSyuSizeArr[2], 10);
        } else if (koSyuSizeArr.length == 4) {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
            size3 = convertStringToIntAndMul(koSyuSizeArr[3], 10);
        } else {
            size1 = convertStringToIntAndMul(koSyuSizeArr[1], 10);
            size2 = convertStringToIntAndMul(koSyuSizeArr[0], 10);
        }
    }

    /**
     * phân tích tính vật liệu của vật liệu đang xét và gán vào map thông tin
     *
     * @param kakuKakou mảng chứa các tính vật liệu của vật liệu đang xét
     * @return map các đoạn tính vật liệu chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
     * còn value của kaKouPairs cũng là mảng chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
     */
    private static Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> getToriaiData(String[] kakuKakou) throws TimeoutException {
        //reset lại map
        seiHinMap.clear();

        rowToriAiNum = 0;
        // tạo map
        Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = new LinkedHashMap<>();

        // nếu không có thông tin thì thoát
        if (kakuKakou == null) {
            return kaKouPairs;
        }

        // lặp qua các đoạn bozai và thêm chúng vào map chứa toàn bộ thông tin vật liệu, tính từ 1 vì 0 là phần chứa thông tin vật liệu
        for (int i = 1; i < kakuKakou.length; i++) {
            // lấy kirirosu tại lần 1
            if (i == 1) {
                kirirosu = extractValue(kakuKakou[i], "切りﾛｽ設定:", "mm");
                if (!kirirosu.equalsIgnoreCase("2.0")) {
                    kirirosu20 = false;
                }
            }

            // lấy đoạn bozai đang lặp
            String kaKouText = kakuKakou[i];

            // map chứa cặp chiều dài, số lượng bozai
            Map<StringBuilder, Integer> kouZaiChouPairs = new LinkedHashMap<>();
            // map chứa cặp key là mảng chứa tên + chiều dài sản phẩm, value là số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = new LinkedHashMap<>();

            // tạo mảng chứa các dòng trong đoạn bozai
            String[] kaKouLines = kaKouText.split("\n");

            // duyệt qua các dòng để thêm vào map
            for (String line : kaKouLines) {
                // nếu dòng có 鋼材長 và 本数 thì là dòng chứa bozai
                // lấy bozai và số lượng thêm vào map
                // mẫu định dạng "#.##". Mẫu này chỉ hiển thị phần thập phân nếu có, và tối đa là 2 chữ số thập phân.

                DecimalFormat df = new DecimalFormat("#.##");
                if (line.contains("鋼材長:") && line.contains("本数:")) {
                    String kouZaiChou = extractValue(line, "鋼材長:", "mm").trim();
                    String kouZaiHonSuu = extractValue(line, "本数:", " ").split(" ")[0].trim();

                    kouZaiChouPairs.put(new StringBuilder().append(df.format(Double.parseDouble(kouZaiChou))), convertStringToIntAndMul(kouZaiHonSuu, 1));
                }

                // nếu dòng chứa 名称 thì là dòng sản phẩm
                if (line.contains("名称")) {
                    // lấy vùng chứa tên và chiều dài sản phẩm
                    String meiSyouLength = extractValue(line, "名称", "mm x").trim();
                    // tách vùng trên thành mảng chứa các phần tử tên và chiều dài
                    String[] meiSyouLengths = meiSyouLength.split(" ");

                    // tạo biến chứa tên
                    String name = "";
                    // vì vùng chứa chiều dài có thể có dấu cách nên phải lấy từ phần tử đầu tiên đến phần tử trước phần tử cuối cùng
                    // và cuối tên sẽ không thêm dấu cách
                    for (int j = 0; j < meiSyouLengths.length - 1; j++) {
                        String namej = meiSyouLengths[j];
                        name = name.concat(namej + " ");
                    }
                    // xóa dấu cách ở 2 đầu
                    name = name.trim();

                    // lấy vùng chứa chiều dài là vùng cuối cùng trong mảng tên
                    String length = meiSyouLengths[meiSyouLengths.length - 1].trim();

                    Double dLength = Double.parseDouble(length);

                    // thêm tên và chiều dài vào mảng, tên với ứng dụng này thì không cần
                    StringBuilder[] nameAndLength = {new StringBuilder(), new StringBuilder().append(df.format(dLength))};

                    // lấy số lượng sản phẩm
                    String meiSyouHonSuu = extractValue(line, "mm x", "本").trim();
                    int honSuu = convertStringToIntAndMul(meiSyouHonSuu, 1);
                    // tổng số lượng sản phẩm trong bozai đang duyệt, nó là số lượng của nó x số lượng bozai
                    int totalHonSuu = Integer.parseInt(extractValue(line, "本(", "本)(").trim());

                    // nếu sản phẩm đã có trong map thì lấy số lượng trong map rồi xóa sản phẩm đi rồi thêm lại sản phẩm với
                    // số lượng trong map đã lấy + số lượng hiện tại
                    // nếu chưa có trong map thì thêm sản phẩm với số lượng hiện tại
                    if (seiHinMap.get(dLength) != null) {
                        int oldNum = seiHinMap.get(dLength);
                        seiHinMap.remove(dLength);
                        seiHinMap.put(dLength, totalHonSuu + oldNum);
                    } else {
                        seiHinMap.put(dLength, totalHonSuu);
                    }

                    // thêm cặp tên + chiều dài và số lượng vào map
                    meiSyouPairs.put(nameAndLength, honSuu);
                }
            }

            // thêm 2 map chứa thông tin vật liệu vào map gốc
            kaKouPairs.put(kouZaiChouPairs, meiSyouPairs);
        }

        // cho các key của map đã lấy được vào list các chiều dài + số lượng của vật liệu đang tính và xắp xếp nó để list sẽ hiển thị trong excel
        seiHinList.setAll(seiHinMap.keySet());
        seiHinList.sort((o1, o2) -> {
            return o1.compareTo(o2);
        });


//        // in thông tin vật liệu
//        kaKouPairs.forEach((kouZaiChouPairs, meiSyouPairs) -> {
//            kouZaiChouPairs.forEach((key, value) -> System.out.println("\n" + key.toString() + " : " + value));
//            meiSyouPairs.forEach((key, value) -> System.out.println(key[0].toString() + " " + key[1].toString() + " : " + value));
//        });

        // lặp qua các phần tử của map kaKouPairs để tính số dòng sản phẩm đã lấy được
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> e : kaKouPairs.entrySet()) {

            // lấy map chiều dài bozai và số lượng
            Map<StringBuilder, Integer> kouZaiChouPairs = e.getKey();
            // lấy map tên + chiều dài sản phẩm và số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = e.getValue();
            // tạo biến chứa số lượng bozai
            int kouZaiNum = 1;
            // lặp qua map bozai lấy giá trị số lượng bozai
            for (Map.Entry<StringBuilder, Integer> entry : kouZaiChouPairs.entrySet()) {
                kouZaiNum = entry.getValue();
            }

            // lấy kết quả số dòng sản phẩm đã lấy được bằng cách lấy số dòng của các lần lặp trước + số dòng của lần này(kouZaiNum * meiSyouPairs.size())
            // meiSyouPairs.size chính là số sản phẩm của bozai đang lặp
            rowToriAiNum += kouZaiNum * meiSyouPairs.size();
        }

        // đoạn này chỉ dùng cho tạo file chl
        /*// nếu số dòng lớn hơn 99 th cho bằng 99 rồi ném ngoại lệ timeout để cho chương trình biết rồi hiển thị thông báo
        if (rowToriAiNum > 99) {
            rowToriAiNum = 99;
            System.out.println("vượt quá 99 hàng");
            // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2 để in ra thông báo
            fileName = fileExcelName + " " + kouSyu + ".sysc2";
            throw new TimeoutException();
        }*/

//        System.out.println(rowToriAiNum);
        System.out.println("\n" + kirirosu);

        // trả về map kết quả để ghi vào file excel
        return kaKouPairs;
    }

    /**
     * chia file theo số dòng sản phẩm
     * nếu vượt quá 99 dòng sẽ chia thành 2 file, file 1 là dưới 99 dòng rồi thêm nó vào list file, file 2 là phần còn lại
     * tiếp tục gọi lại hàm(đệ quy) và truyền file 2 vào và thực hiện tương tự đến khi không cần chia thành file 2 nữa tức
     * file 2 có size = 0 thì dừng
     *
     * @param kaKouPairs map chứa tính vật liệu đại diện cho file đang gọi
     */
    private static void divFile(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs) throws TimeoutException {

        Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> map1 = new LinkedHashMap<>();
        Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> map2 = new LinkedHashMap<>();

        int numLoop1 = 0;
        int numRow = 0;
        // lặp qua các phần tử của map kaKouPairs để tính số dòng sản phẩm đã lấy được
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> e : kaKouPairs.entrySet()) {
            numLoop1 += 1;
            // lấy map chiều dài bozai và số lượng
            Map<StringBuilder, Integer> kouZaiChouPairs = e.getKey();
            // lấy map tên + chiều dài sản phẩm và số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = e.getValue();

            // nếu số sản phẩm trong 1 bozai  lớn hơn 99 thì không
            // thể chia sang file khác được và phải báo lỗi
            if (meiSyouPairs.size() > 99) {
                System.out.println("vượt quá 99 hàng");
                // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2 để in ra thông báo
                fileName = fileChlName + " " + kouSyu + ".sysc2";
                throw new TimeoutException();
            }

            // tạo biến chứa chiều dài và số lượng bozai
            StringBuilder koZaiLength = new StringBuilder();
            int kouZaiNum = 1;
            // lặp qua map bozai lấy giá trị số lượng bozai
            for (Map.Entry<StringBuilder, Integer> entry : kouZaiChouPairs.entrySet()) {
                koZaiLength = entry.getKey();
                kouZaiNum = entry.getValue();
            }

            // thêm toriai vào map 1 đến số lượng dòng <= 99

            // biến nhớ vượt quá 100 dòng
            boolean is100 = false;

            // biến nhớ đã lặp qua kouZaiNum bao nhiêu lần mà số dòng vẫn chưa vượt quá 99
            int numKouZaiChouMap1 = 0;

            for (int i = 1; i <= kouZaiNum; i++) {
                // nếu số dòng tính trước trong lần này vượt quá 99 dòng thì lấy biến nhớ số lần lặp hợp lệ
                // cho biến nhớ quá 100 là true và thoát lặp
                if (numRow + meiSyouPairs.size() > 99) {

                    is100 = true;
                    break;
                }
                // lấy biến nhớ số lần lặp đã hợp lệ là số dòng chưa vượt qua 99
                numKouZaiChouMap1 = i;
                // lấy kết quả số dòng sản phẩm đã lấy được bằng cách lấy số dòng của các lần lặp trước + số dòng của lần này(numRow += meiSyouPairs.size())
                // meiSyouPairs.size chính là số sản phẩm của bozai đang lặp
                numRow += meiSyouPairs.size();
            }

            Map<StringBuilder, Integer> newKouZaiChouPairs = new HashMap<>();
            // nếu số dòng hợp lệ > 0 thì tức là có dòng hợp lệ
            if (numKouZaiChouMap1 > 0) {
                // tạo map của chiều dài bozai và số lượng mới với số lượng là số lần lặp hợp lệ của kouZaiNum mà chưa vượt quá 99 dòng
                // rồi thêm map mới này vào map chứa toriai
                newKouZaiChouPairs = new HashMap<>();
                newKouZaiChouPairs.put(koZaiLength, numKouZaiChouMap1);
                // thêm map mới vào map chứa toriai là map1
                map1.put(newKouZaiChouPairs, meiSyouPairs);
            }

            // nếu số lần lặp của map 1 numKouZaiChouMap1 < kouZaiNum tức là nó chưa đi hết kouZaiNum mà đã vượt quá 99 dòng
            // thêm số lượng còn lại vào map2
            if (numKouZaiChouMap1 < kouZaiNum) {
                newKouZaiChouPairs = new HashMap<>();
                newKouZaiChouPairs.put(koZaiLength, kouZaiNum - numKouZaiChouMap1);
                // thêm map mới vào map chứa toriai là map2
                map2.put(newKouZaiChouPairs, meiSyouPairs);
            }

            // nếu lần lặp của bozai này đã quá 100 dòng thì thoát để tạo vòng lặp tiếp thêm vào map 2
            if (is100) {
                break;
            }

        }

        // nếu lần lặp 1 vẫn chưa lặp hết số bozai thì lặp nốt phần còn lại cho vào map 2
        if (numLoop1 < kaKouPairs.size() - 1) {
            int numLoop2 = 0;
            // lặp qua các phần tử của map kaKouPairs để tính số dòng sản phẩm đã lấy được
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> e : kaKouPairs.entrySet()) {
                numLoop2 += 1;

                // nếu numLoop2 <= numLoop1 thì tức là lần lặp này vẫn thuộc lần lặp của map1 đã lặp ở bên trên
                if (numLoop2 <= numLoop1) {
                    continue;
                }

                // lấy map chiều dài bozai và số lượng
                Map<StringBuilder, Integer> kouZaiChouPairs = e.getKey();
                // lấy map tên + chiều dài sản phẩm và số lượng
                Map<StringBuilder[], Integer> meiSyouPairs = e.getValue();

                // thêm toriai vào map 2
                map2.put(kouZaiChouPairs, meiSyouPairs);
            }
        }

        // thêm map1 vào vào list file
        fileList.add(map1);

        // nếu map2 không có phần tử nào tức đã hoàn thành chia file
        if (map2.size() == 0) {
            return;
        }

        // gọi đệ quy hàm chia file và truyền map2 vào để tiếp tục chia map 2
        divFile(map2);

    }

    private static int checkRowNum(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs) throws TimeoutException {
        rowToriAiNum = 0;

        // lặp qua các phần tử của map kaKouPairs để tính số dòng sản phẩm đã lấy được
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> e : kaKouPairs.entrySet()) {

            // lấy map chiều dài bozai và số lượng
            Map<StringBuilder, Integer> kouZaiChouPairs = e.getKey();
            // lấy map tên + chiều dài sản phẩm và số lượng
            Map<StringBuilder[], Integer> meiSyouPairs = e.getValue();

            // nếu số sản phẩm trong 1 bozai  lớn hơn 99 thì không
            // thể chia sang file khác được và phải báo lỗi
            if (meiSyouPairs.size() > 99) {
                System.out.println("vượt quá 99 hàng");
                // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2 để in ra thông báo
                fileName = fileChlName + " " + kouSyu + ".sysc2";
                throw new TimeoutException();
            }

            // tạo biến chứa số lượng bozai
            int kouZaiNum = 1;
            // lặp qua map bozai lấy giá trị số lượng bozai
            for (Map.Entry<StringBuilder, Integer> entry : kouZaiChouPairs.entrySet()) {
                kouZaiNum = entry.getValue();
            }

            // lấy kết quả số dòng sản phẩm đã lấy được bằng cách lấy số dòng của các lần lặp trước + số dòng của lần này(kouZaiNum * meiSyouPairs.size())
            // meiSyouPairs.size chính là số sản phẩm của bozai đang lặp
            rowToriAiNum += kouZaiNum * meiSyouPairs.size();
        }

        return rowToriAiNum;
    }

    private static void writeDataToExcel(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<ExcelFile> excelFileNames) throws FileNotFoundException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Ghi thời gian hiện tại vào ô A1
        Row row1 = sheet.createRow(0);
        Cell cellA1 = row1.createCell(0);

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmm");
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

        // Tăng thời gian lên timePlus phút
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(currentDate);

        cellA1.setCellValue(newTime + "+" + timePlus);

        // Ghi size1, size2, size3, 1 vào ô A2, B2, C2, D2
        Row row2 = sheet.createRow(1);
        row2.createCell(0).setCellValue(size1);
        row2.createCell(1).setCellValue(size2);
        row2.createCell(2).setCellValue(size3);
        row2.createCell(3).setCellValue(1);

        // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào ô A3, B3, C3, D3
        Row row3 = sheet.createRow(2);
        row3.createCell(0).setCellValue(koSyuNumMark);
        row3.createCell(1).setCellValue(1);
        row3.createCell(2).setCellValue(rowToriAiNum);
        row3.createCell(3).setCellValue(1);

        int rowIndex = 3;

        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;
        // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
            if (rowIndex >= 102) break;

            Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
            Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

            String keyTemp = "";
            int valueTemp = 0;

            // Ghi dữ liệu từ mapkey vào ô D4
            for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                keyTemp = String.valueOf(kouZaiEntry.getKey());
                valueTemp = kouZaiEntry.getValue();
                // cộng thêm chiều dài của bozai * số lượng vào tổng
                kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
            }

            // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
            for (int i = 0; i < valueTemp; i++) {
                int j = 0;
                for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                    if (rowIndex >= 102) break;
                    // chiều dài sản phẩm
                    String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                    // số lượng sản phẩm
                    String num = meiSyouEntry.getValue().toString();

                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(leng);
                    row.createCell(1).setCellValue(num);
                    row.createCell(2).setCellValue(String.valueOf(meiSyouEntry.getKey()[0]));

                    // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                    seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);
                    j++;
                }
                sheet.getRow(rowIndex - j).createCell(3).setCellValue(keyTemp);
            }
        }

/*        // không cần tạo nữa vì chiều dài bozai sẽ ghi vào cột 4
        // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
        // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
        // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
        for (int i = 0; i < 5; i++) {
            if (rowIndex <= i + 3) {
                sheet.createRow(i + 3);
            }
        }

        // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + kouSyu vào ô D4, D5, D6, D7
        sheet.getRow(3).createCell(3).setCellValue(kouJiMe);
        sheet.getRow(4).createCell(3).setCellValue(kyakuSakiMei);
        sheet.getRow(5).createCell(3).setCellValue(shortNouKi);
        sheet.getRow(6).createCell(3).setCellValue(kirirosu);
        sheet.getRow(7).createCell(3).setCellValue(fileChlName + " " + kouSyu);*/

        // Ghi giá trị 0 vào các ô A99, B99, C99, D99
        Row lastRow = sheet.createRow(rowIndex);
        lastRow.createCell(0).setCellValue(0);
        lastRow.createCell(1).setCellValue(0);
        lastRow.createCell(2).setCellValue(0);
        lastRow.createCell(3).setCellValue(0);

        String[] linkarr = pdfPath.split("\\\\");
//        fileName = linkarr[linkarr.length - 1].split("\\.")[0] + " " + kouSyu + ".xlsx";
        fileName = fileChlName + " " + kouSyu + ".xlsx";
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";
        String excelPath = excelDirPath + "\\" + fileName;

        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(excelPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream(excelPath)) {
            workbook.write(fileOut);
            workbook.close();
        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(excelPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
                System.out.println("File is set to read-only.");
            } else {
                System.out.println("Failed to set file to read-only.");
            }
        } else {
            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        excelFileNames.add(new ExcelFile(fileName, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    private static void writeDataToCSV(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<ExcelFile> excelFileNames) throws FileNotFoundException {

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmm");
//        // Tạo thêm fomat có thêm giây
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

        /*// Tăng thời gian lên timePlus phút
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(newDate);*/

        // lấy thời gian hiện tại với fomat đã chọn
        String currentTime = sdf.format(currentDate);

        String[] linkarr = pdfPath.split("\\\\");
//        fileName = linkarr[linkarr.length - 1].split("\\.")[0] + " " + kouSyu + ".csv";
        fileName = fileChlName + " " + kouSyu + ".csv";
//        // tạo tên file có gắn thêm thời gian để không trùng với file trước đó
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";
        String csvPath = excelDirPath + "\\" + fileName;
        System.out.println("dir path: " + excelDirPath);
        System.out.println("filename: " + fileName);

        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(csvPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        } else {
//            System.out.println("File không tồn tại.");
        }
        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;
        try (CSVWriter writer = new CSVWriter(new OutputStreamWriter(new FileOutputStream(csvPath), Charset.forName("MS932")))) {


            writer.writeNext(new String[]{currentTime + "+" + timePlus});

            // Ghi size1, size2, size3, 1 vào dòng tiếp theo
            writer.writeNext(new String[]{String.valueOf(size1), String.valueOf(size2), String.valueOf(size3), "1"});

            // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào dòng tiếp theo
            writer.writeNext(new String[]{koSyuNumMark, "1", String.valueOf(rowToriAiNum), "1"});

            List<String[]> toriaiDatas = new LinkedList<>();

            int rowIndex = 3;

            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                String keyTemp = "";
                int valueTemp = 0;

                // Ghi dữ liệu từ mapkey vào ô D4
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();

                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukei += Double.parseDouble(keyTemp) * valueTemp;
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0;
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;

                        String[] line = new String[4];
                        rowIndex++;

                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();
                        // ghi chiều dài sản phẩm
                        line[0] = leng;
                        // ghi số lượng sản phẩm
                        line[1] = num;
                        line[2] = String.valueOf(meiSyouEntry.getKey()[0]);

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukei += Double.parseDouble(leng) * Double.parseDouble(num);
                        toriaiDatas.add(line);
                        j++;
                    }
                    toriaiDatas.get(toriaiDatas.size() - j)[3] = keyTemp;
                }
            }

/*            // không cần tạo nữa vì chiều dài bozai sẽ ghi vào cột 4
            // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
            // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
            // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
            for (int i = 0; i < 5; i++) {
                if (rowIndex <= i + 3) {
                    toriaiDatas.add(new String[4]);
                }
            }

            // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + " " + kouSyu vào ô D4, D5, D6, D7
            toriaiDatas.get(0)[3] = kouJiMe;
            toriaiDatas.get(1)[3] = kyakuSakiMei;
            toriaiDatas.get(2)[3] = shortNouKi;
            toriaiDatas.get(3)[3] = kirirosu;
            toriaiDatas.get(4)[3] = fileChlName + " " + kouSyu;*/

            writer.writeAll(toriaiDatas);

            // Ghi giá trị 0 vào các ô A99, B99, C99, D99
            writer.writeNext(new String[]{"0", "0", "0", "0"});


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(csvPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
//                System.out.println("File is set to read-only.");
            } else {
//                System.out.println("Failed to set file to read-only.");
            }
        } else {
//            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        excelFileNames.add(new ExcelFile(fileName, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    /**
     * ghi tính vật liệu của vật liệu đang xét trong map vào file mới
     *
     * @param kaKouPairs     map chứa tính vật liệu
     * @param timePlus       thời gian hoặc chỉ số cộng thêm vào ô time để tránh bị trùng tên  time giữa các file
     * @param excelFileNames list chứa danh sách các file đã tạo
     * @param fileListSize
     * @param k
     */
    private static void writeDataToChl(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int timePlus, ObservableList<ExcelFile> excelFileNames, int fileListSize, int k) throws FileNotFoundException {

        // Ghi thời gian hiện tại vào dòng đầu tiên
        Date currentDate = new Date();
        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
        SimpleDateFormat sdfMiliS = new SimpleDateFormat("SSS");
//        // Tạo thêm fomat có thêm giây
//        SimpleDateFormat sdfSecond = new SimpleDateFormat("yyMMddHHmmss");

/*        // Tăng thời gian lên timePlus phút
        // hiện tại không dùng đoạn code này nữa
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        calendar.add(Calendar.MINUTE, timePlus);

        // Lấy thời gian sau khi tăng
        Date newDate = calendar.getTime();

        String newTime = sdf.format(newDate);*/

        // lấy thời gian hiện tại với fomat đã chọn
        String currentTime = sdf.format(currentDate);
        String currentTimeMiliS = String.valueOf((Integer.parseInt(sdfMiliS.format(currentDate)) / 100));
        currentTime = currentTime + currentTimeMiliS;

        fileName = fileChlName + " " + kouSyu;
        // nếu danh sách file của vật liệu này nhiều hơn 1 thì thêm hậu tố chỉ số thứ tự của file
        if (fileListSize > 1) {
            fileName = fileName + "(" + k + ")";
        }
        // lấy tên file chl trong tiêu đề gắn thêm tên vật liệu + .sysc2
        fileName = fileName + ".sysc2";

//        // tạo tên file có gắn thêm thời gian để không trùng với file trước đó
//        String fileNameAndTime = linkarr[linkarr.length - 1].split("\\.")[0] + "(" + sdfSecond.format(currentDate) + ")--" + kouSyu + ".csv";

        String chlPath = chlDirPath + "\\" + fileName;
        System.out.println("dir path: " + excelDirPath);
        System.out.println("filename: " + fileName);


        // Tạo đối tượng File đại diện cho file cần xóa
        File file = new File(chlPath);

        // Kiểm tra nếu file tồn tại và xóa nó
        // vì nếu file đang được mở thì không thể ghi đè nhưng do file là readonly nên có thể xóa dù đang mở
        // xóa xong file thì có thể ghi lại file mới mà không bị lỗi không thể ghi đè
        if (file.exists()) {
            if (file.delete()) {
//                System.out.println("File đã được xóa thành công.");
            } else {
//                System.out.println("Xóa file thất bại.");
            }
        } else {
//            System.out.println("File không tồn tại.");
        }

        // tổng chiều dài các kozai
        double kouzaiChouGoukeiTempt = 0;
        double seiHinChouGoukeiTempt = 0;
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(chlPath, Charset.forName("MS932")))) {

            writer.write(currentTime + timePlus + ",,,");
            writer.newLine();


            // Ghi size1, size2, size3, 1 vào dòng tiếp theo
            writer.write(size1 + "," + size2 + "," + size3 + "," + "1");
            writer.newLine();

            // Ghi koSyuNumMark, 1, 99, 1 vào dòng tiếp theo, rowToriAiNum sẽ được sử dụng sau khi ước tính ghi đến hàng 102
            writer.write(koSyuNumMark + "," + "1" + "," + "99" + "," + "1");
            writer.newLine();

            // tạo list chứa các mảng, mỗi mảng là 1 dòng cần ghi theo fomat của chl
            List<String[]> toriaiDatas = new LinkedList<>();

            int rowIndex = 3;

            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            // kaKouPairs là map chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai
            // còn value của kaKouPairs cũng là map chứa các cặp key là tên + chiều dài sản phẩm, value là số lượng sản phẩm
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                // chiều dài bozai
                String keyTemp = "";
                // số lượng bozai
                int valueTemp = 0;


                // Ghi dữ liệu bozai từ mapkey vào ô D4 kouZaiChouPairs
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {
                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();
                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukeiTempt += Double.parseDouble(keyTemp) * valueTemp;
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                // số lượng bozai là bao nhiêu thì phải ghi bấy nhiêu lần
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0; // đếm số hàng đã ghi
                    // lặp qua map sản phẩm, tính chiều dài map bằng j
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;

                        // tạo mảng lưu dòng đang lặp gồm 4 phần tử lần lượt là
                        // chiều dài sản phẩm, số lượng sản phẩm, tên sản phẩm, chiều dài bozai
                        String[] line = new String[4];
                        rowIndex++;

                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();
                        // ghi chiều dài sản phẩm
                        line[0] = leng;
                        // ghi số lượng sản phẩm
                        line[1] = num;
                        // ghi tên sản phẩm
                        line[2] = String.valueOf(meiSyouEntry.getKey()[0]);
                        // ghi vào phần tử thứ 3 của mảng giá trị rỗng để tránh giá trị null
                        line[3] = "";

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukeiTempt += Double.parseDouble(leng) * Double.parseDouble(num);

                        // thêm hàng sản phẩm vừa tạo vào list
                        toriaiDatas.add(line);
                        // tăng số hàng lên 1
                        j++;
                    }
                    // ghi vào cột 4 ([3]) chiều dài bozai khi ghi xong 1 lượt sản phẩm + số lượng
                    // tính vị trí của nó bằng cách lấy size của list kaKouPairs - chiều dài map sản phẩm
                    toriaiDatas.get(toriaiDatas.size() - j)[3] = keyTemp;
                }
            }

/*
            // nếu không có hàng sản phẩm nào thì sẽ chưa tạo hàng 4, 5, 6, 7, 8 và rowIndex vẫn là 3
            // cần tạo thêm 4 hàng này để ghi các thông tin kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName bên dưới
            // không cần tạo nữa vì ghi file sysc2 sẽ ghi xuống cuối
            for (int i = 0; i < 5; i++) {
                if (rowIndex <= i + 3) {
                    toriaiDatas.add(new String[4]);
                }
            }
*/

/*
            // Ghi kouJiMe, kyakuSakiMei, shortNouKi, kirirosu, fileName + " " + kouSyu vào ô D4, D5, D6, D7
            // không cần tạo nữa vì ghi file sysc2 sẽ ghi xuống cuối và vì chiều dài bozai sẽ ghi vào cột 4
            // thay vì cột 3 như trước nên không thể ghi thêm các thông tin này vào cột 4 nữa
            toriaiDatas.get(0)[3] = kouJiMe;
            toriaiDatas.get(1)[3] = kyakuSakiMei;
            toriaiDatas.get(2)[3] = shortNouKi;
            toriaiDatas.get(3)[3] = kirirosu;
            toriaiDatas.get(4)[3] = fileChlName + " " + kouSyu;
*/
            // lặp qua list chứa các dòng toriaiDatas
            for (String[] line : toriaiDatas) {

/*                // cách ghi này không dùng được nữa vì cách ghi phần tử cuối cùng đã thay đổi
                for (String length : line) {
                    writer.write(length + ",");
                }*/

                // mỗi dòng là 1 mảng nên lặp qua mảng ghi các phần tử vào dòng phân tách nhau bởi dấu (,)
                for (int i = 0; i < line.length; i++) {
                    if (i == line.length - 1) {
                        writer.write(line[i]);
                    } else {
                        writer.write(line[i] + ",");
                    }
                }
                writer.newLine();
            }

            // ghi nốt các dòng còn lại không có sản phẩn ",,," để đủ 99 sản phẩm
            for (int i = toriaiDatas.size(); i < 99; i++) {
                writer.write(",,,");
                writer.newLine();
            }


            // Ghi giá trị 0 vào dòng tiếp theo là dòng 103
            writer.write("0,0,0,0");
            writer.newLine();
            // ghi 20 và kirirosu vào dòng tiếp
            writer.write("20.0," + kirirosu + ",,");
            writer.newLine();
            // ghi các tên và ngày vào dòng tiếp
            writer.write(kouJiMe + "," + kyakuSakiMei + "," + shortNouKi + ",");
            writer.newLine();
            // dòng tiếp theo là ghi 備考１、備考２ theo định dạng 備考１,備考２,, nhưng không có nên không cần chỉ ghi (,,,)
            writer.write(",,,");
            writer.newLine();
            // ghi dấu hiệu nhận biết kết thúc
            writer.write("END,,,");
            writer.newLine();


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }

        // Đặt quyền chỉ đọc cho file
        File readOnly = new File(chlPath);
        if (readOnly.exists()) {
            boolean result = readOnly.setReadOnly();
            if (result) {
//                System.out.println("File is set to read-only.");
            } else {
//                System.out.println("Failed to set file to read-only.");
            }
        } else {
//            System.out.println("File does not exist.");
        }

        System.out.println("tong chieu dai bozai " + kouzaiChouGoukeiTempt);
        System.out.println("tong chieu dai san pham " + seiHinChouGoukeiTempt);

        // cộng thêm chiều các tổng chiều dài file hiện tại vào tổng các chiều dài của các file
        kouzaiChouGoukei += kouzaiChouGoukeiTempt;
        seiHinChouGoukei += seiHinChouGoukeiTempt;

        // nếu đang ghi file cuối cùng của vật liệu thì mới ghi các tổng chiều dài
        if (fileListSize == k) {
            kouzaiChouGoukeiTempt = kouzaiChouGoukei;
            seiHinChouGoukeiTempt = seiHinChouGoukei;
        } else {
            kouzaiChouGoukeiTempt = 0;
            seiHinChouGoukeiTempt = 0;
        }

        // thêm file vào list hiển thị
        excelFileNames.add(new ExcelFile(fileName, kouSyuName, kouzaiChouGoukeiTempt, seiHinChouGoukeiTempt));

    }

    /**
     * trả về đoạn text nằm giữa startDelimiter và endDelimiter
     *
     * @param text           đoạn văn bản chứa thông tin tìm kiếm
     * @param startDelimiter đoạn text phía trước vùng cần tìm
     * @param endDelimiter   đoạn text phía sau vùng cần tìm
     * @return đoạn text nằm giữa startDelimiter và endDelimiter
     */
    private static String extractValue(String text, String startDelimiter, String endDelimiter) {
        // lấy index của startDelimiter + độ dài của nó để bỏ qua nó và xác định được index bắt đầu của đoạn text nó bao ngoài, chính là đoạn text cần tìm
        int startIndex = text.indexOf(startDelimiter) + startDelimiter.length();
        // lấy index của endDelimiter bắt đầu tìm từ index của startDelimiter để tránh tìm kiếm trong các vùng khác phía trước không liên quan, đây chính là
        // index cuối cùng của đoạn text cần tìm
        int endIndex = text.indexOf(endDelimiter, startIndex);
//        System.out.println(text);
        // trả về đoạn text cần tìm bằng 2 index vừa xác định ở trên
        return text.substring(startIndex, endIndex).trim();
    }


    /**
     * chuyển đổi text nhập vào sang số BigDecimal rồi nhân với hệ số và trả về với kiểu int
     *
     * @param textNum    text cần chuyển
     * @param multiplier hệ số
     * @return số int đã nhân với hệ số
     */
    public static int convertStringToIntAndMul(String textNum, int multiplier) {
        BigDecimal bigDecimalNum = null;
        try {
            bigDecimalNum = new BigDecimal(textNum);
            // nhân số thực num với hệ số truyền vào
            bigDecimalNum = bigDecimalNum.multiply(new BigDecimal(multiplier));

        } catch (NumberFormatException e) {
            System.out.println("Lỗi chuyển đổi chuỗi không phải số thực sang số");
            System.out.println(textNum);

        }
        if (bigDecimalNum != null) {
            // Làm tròn đến số nguyên gần nhất
            return bigDecimalNum.setScale(0, RoundingMode.DOWN).intValueExact();
        }
        return 0;
    }


    /**
     * chuyển các thông số của tính vật liệu trong file 3bc sang excel
     *
     * @param kaKouPairs     thông số chiều dài vật liệu và chiều dài + số lượng của các sản phẩm tính cho cây vật liệu đó
     * @param sheetIndex     thứ tự tạo sheet
     * @param excelFileNames tên file excel
     * @throws FileNotFoundException
     */
    private static void writeDataToExcelToriai(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int sheetIndex, ObservableList<ExcelFile> excelFileNames) throws FileNotFoundException {

        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;

        // tạo luồng đọc ghi file
        try (FileInputStream file = new FileInputStream(excelCopyPath)) {
            Workbook workbook = new XSSFWorkbook(file);

            // nếu tên vật liệu có chứa [ thì phải đổi sang U vì tên này sẽ đặt tên cho sheet nên [ không dùng được
            if (kouSyu.contains("[")) {
                kouSyu = kouSyu.replace("[", "U");
            }

            // Lấy index sheet gốc cần sao chép
            int sheetSampleIndex = 0;
            // sao chép sheet gốc sang một sheet mới
            workbook.cloneSheet(sheetSampleIndex);
            // đổi tên sheet mới theo tên vật liệu đang duyệt, sheetIndex là chỉ số của sheet mới
            workbook.setSheetName(sheetIndex, kouSyu);
            // lấy ra sheet mới
            Sheet sheet = workbook.getSheetAt(sheetIndex);


            Date currentDate = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

            String time = sdf.format(currentDate);
            // Ghi thời gian hiện tại vào ô C1
            sheet.getRow(0).getCell(2).setCellValue(time);

            // Ghi tên khách hàng vào ô G6
            sheet.getRow(0).getCell(6).setCellValue(kyakuSakiMei);

            // Ghi bikou vào ô M12
            sheet.getRow(0).getCell(12).setCellValue(bikou);

            // Ghi shortNouKi vào ô S18
            sheet.getRow(0).getCell(18).setCellValue(shortNouKi);

            // Ghi saizu vào ô C2, chưa dùng
            sheet.getRow(1).getCell(2).setCellValue("");

            // Ghi chuyuBan vào ô I8
            sheet.getRow(1).getCell(8).setCellValue(chuyuBan);
            // Ghi teiHaiSha vào ô O14
            sheet.getRow(1).getCell(14).setCellValue(teiHaiSha);

            // lấy số loại bozai và sản phẩm
            int soBoZai = kaKouPairs.size();
            int soSanPham = seiHinList.size();

            // nếu số bozai nhiều hơn 15 bao nhiêu thì thêm số cột bozai với số lượng đó
            // copy và paste giá trị cho cột mới cho giống giá trị với các cột còn lại
            if (soBoZai > 15) {

                // thêm j lần các cột mới tại các cột công thức
                for (int j = 0; j < soBoZai - 15; j++) {
                    sheet.shiftColumns(4, sheet.getRow(6).getLastCellNum(), 1);
                    sheet.shiftColumns(4 + 23 + j, sheet.getRow(6).getLastCellNum(), 1);
                    sheet.shiftColumns(4 + 41 + 2 * j, sheet.getRow(6).getLastCellNum(), 1);
//                System.out.println("last col: " + sheet.getRow(6).getLastCellNum());

                    // dịch chuyển 3 hàng tiêu đề về vị trí ban đầu sau khi bị dịch chuyển sang phải 1 hàng
                    for (int i = 0; i < 3; i++) {
                        Row row = sheet.getRow(i);
                        row.shiftCellsLeft(5, 10000, 1);
                    }

                    // sửa lại công thức tất cả các ô có giá trị L về K vì sau khi dịch chuyển 3 hàng tiêu đề về vị trí ban đầu
                    // công thức bị sai
                    for (int i = 26 + j; i <= 41 + 2 * j; i++) {
                        // row index 6 tức là hàng 7 vì ban đầu chỉ có 1 hàng 7 có công thức do chưa thêm các hàng mới
                        Row row = sheet.getRow(6);
                        Cell cell = row.getCell(i);

                        if (cell != null && cell.getCellType() == CellType.FORMULA) {
                            String formula = cell.getCellFormula();
                            formula = formula.replaceAll("\\$L\\$3", "\\$K\\$3");
                            cell.setCellFormula(formula);
                        }
                    }

                    Cell srcCell;
                    Cell destCell;

                    // sao chép ô từ cột 3 sang cột 4 từ hàng 3 đến hàng 9 trong 2 cột này
                    // cần tạo cell ở cột 4 bị phép dịch chuyển cột ở trên thực chất chưa tạo cell mới
                    for (int i = 3; i <= 9; i++) {
                        Row row = sheet.getRow(i);
                        // Sao chép ô từ cột srcColumn sang destColumn
                        srcCell = row.getCell(3);
                        destCell = row.createCell(4);
                        copyCellWithFormulaUpdate(srcCell, destCell, 1);
                    }

                    // tại hàng 7 copy ô từ cột 26 sang 27
                    Row row7Formula = sheet.getRow(6);
                    srcCell = row7Formula.getCell(26 + j);
                    destCell = row7Formula.createCell(27 + j);
                    copyCellWithFormulaUpdate(srcCell, destCell, 1);

                    // tại hàng 7 copy ô từ cột 44 sang 45
                    srcCell = row7Formula.getCell(44 + 2 * j);
                    destCell = row7Formula.createCell(45 + 2 * j);
                    copyCellWithFormulaUpdate(srcCell, destCell, 1);

                    // tại hàng 4 copy ô từ cột 44 sang 45
                    Row row4Formula = sheet.getRow(3);
                    srcCell = row4Formula.getCell(44 + 2 * j);
                    destCell = row4Formula.createCell(45 + 2 * j);
                    copyCellWithFormulaUpdate(srcCell, destCell, 1);
                }

            }


            // nếu số sản phẩm lớn hơn 1 bao nhiêu lần thì thêm số hàng sản phẩm số lần tương tự
            if (soSanPham > 1) {
                for (int j = 0; j < soSanPham - 1; j++) {
                    // đẩy tất cả các hàng ở dưới hàng index 6 xuống 1 hàng để thừa ra hàng index 7 nhưng nó thực tế vẫn chưa được tạo
                    // sau đó mới tạo hàng index 7
                    sheet.shiftRows(7, sheet.getLastRowNum(), 1);
                    Row srcRow = sheet.getRow(6);
                    // tạo hàng index 7
                    Row destRow = sheet.createRow(7);

                    // Sao chép từng cell từ hàng nguồn sang hàng đích
                    for (int i = 0; i < srcRow.getLastCellNum(); i++) {
                        Cell sourceCell = srcRow.getCell(i);
                        Cell targetCell = destRow.createCell(i);

                        if (sourceCell != null) {
                            copyRowCellWithFormulaUpdate(sourceCell, targetCell, 1);
                        }
                    }
                }
            }

            // ghi tất cả sản phẩm vào excel
            for (int i = 0; i < soSanPham; i++) {
                Double length = seiHinList.get(i);
                // ghi chiều dài sản phẩm
                sheet.getRow(i + 6).getCell(0).setCellValue(length);
                // ghi số lượng sản phẩm
                sheet.getRow(i + 6).getCell(1).setCellValue(seiHinMap.get(length));
            }


            // ghi bozai và sản phẩm trong bozai
            // thể hiện index cột bozai đang thực thi
            int numBozai = 0;
            // lặp qua các cặp tính vật liệu, mỗi cặp gồm "bozai-số lượng trong map kouZaiChouPairs(nằm trong map nhưng thực tế nó chỉ có 1 cặp)"
            // và "các bộ chiều dài sản phẩm và số lượng trong map meiSyouPairs"
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                // Ghi bozai và số lượng của nó
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                    sheet.getRow(3).getCell(3 + numBozai).setCellValue(String.valueOf(kouZaiEntry.getKey()));
                    sheet.getRow(4).getCell(3 + numBozai).setCellValue(String.valueOf(kouZaiEntry.getValue()));

                    kouzaiChouGoukei += Double.parseDouble(String.valueOf(kouZaiEntry.getKey())) * kouZaiEntry.getValue();
                }

                // lặp qua các chiều dài sản phẩm trong cặp tính vật liệu này và tìm trong hàng có chiều dài sản phẩm
                // tương ứng trong cột sản phẩm, dóng sang cột bozai đang tạo là tìm được ô cần ghi sô lượng, sau đó ghi cộng dồn số lượng vào ô đó
                for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                    // chiều dài sản phẩm
                    Double length = Double.valueOf(meiSyouEntry.getKey()[1].toString());
                    // số lượng sản phẩm
                    int num = Integer.parseInt(meiSyouEntry.getValue().toString());

                    // hàng chứa sản phẩm, +6 vì cột chứa sản phẩm bắt đầu chứa các sản phẩm từ hàng thứ 6
                    int indexSeiHinRow = seiHinList.indexOf(length) + 6;

                    // lấy cell chứa số lượng của sản phẩm
                    Cell cellSoLuong = sheet.getRow(indexSeiHinRow).getCell(3 + numBozai);

                    // lấy số lượng cũ của cell
                    double oldNum = 0d;
                    // nếu cell có type là số thì nó đã có số lượng từ trước thì gán nó cho số lượng cũ
                    if (cellSoLuong.getCellType() == CellType.NUMERIC) {
                        oldNum = cellSoLuong.getNumericCellValue();
                    }

                    // nếu số lượng cũ > 0 thì ghi giá trị cell với số lượng cũ + số lượng hiện tại
                    // không thì ghi cell với số lượng hiện tại
                    if (oldNum > 0d) {
                        sheet.getRow(indexSeiHinRow).getCell(3 + numBozai).setCellValue(num + oldNum);
                    } else {
                        sheet.getRow(indexSeiHinRow).getCell(3 + numBozai).setCellValue(num);
                    }

                    // thống kê phục vụ cho hiển thị thông tin trên phầm mềm
                    double totalLength = Double.parseDouble(String.valueOf(meiSyouEntry.getKey()[1])) * Double.parseDouble(meiSyouEntry.getValue().toString());
                    Cell cellSoLuongBozai = sheet.getRow(4).getCell(3 + numBozai);
                    if (cellSoLuongBozai.getCellType() == CellType.STRING) {
                        totalLength *= Double.parseDouble(cellSoLuongBozai.getStringCellValue());
                    }

                    seiHinChouGoukei += totalLength;

                }

                numBozai++;
            }

            /*
            // Ghi koSyuNumMark, 1, rowToriAiNum, 1 vào ô A3, B3, C3, D3
            Row row3 = sheet.createRow(2);
            row3.createCell(0).setCellValue(koSyuNumMark);
            row3.createCell(1).setCellValue(1);
            row3.createCell(2).setCellValue(rowToriAiNum);
            row3.createCell(3).setCellValue(1);

            int rowIndex = 3;

            // tổng chiều dài các kozai
            double kouzaiChouGoukei = 0;
            double seiHinChouGoukei = 0;
            // Ghi dữ liệu từ KA_KOU_PAIRS vào các ô
            for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
                if (rowIndex >= 102) break;

                Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
                Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

                String keyTemp = "";
                int valueTemp = 0;

                // Ghi dữ liệu từ mapkey vào ô D4
                for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {

                    keyTemp = String.valueOf(kouZaiEntry.getKey());
                    valueTemp = kouZaiEntry.getValue();
                    // cộng thêm chiều dài của bozai * số lượng vào tổng
                    kouzaiChouGoukei += Double.parseDouble(String.valueOf(kouZaiEntry.getKey())) * kouZaiEntry.getValue();
                }

                // Ghi dữ liệu từ mapvalue vào ô A4, B4 và các hàng tiếp theo
                for (int i = 0; i < valueTemp; i++) {
                    int j = 0;
                    for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                        if (rowIndex >= 102) break;
                        // chiều dài sản phẩm
                        String leng = String.valueOf(meiSyouEntry.getKey()[1]);
                        // số lượng sản phẩm
                        String num = meiSyouEntry.getValue().toString();

                        Row row = sheet.createRow(rowIndex++);
                        row.createCell(0).setCellValue(leng);
                        row.createCell(1).setCellValue(num);
                        row.createCell(2).setCellValue(String.valueOf(meiSyouEntry.getKey()[0]));

                        // cộng thêm vào chiều dài của sản phẩm * số lượng vào tổng
                        seiHinChouGoukei += Double.parseDouble(String.valueOf(meiSyouEntry.getKey()[1]);) * Double.parseDouble(meiSyouEntry.getValue().toString());
                        j++;
                    }
                    sheet.getRow(rowIndex - j).createCell(3).setCellValue(keyTemp);
                }
            }*/

/*            Sheet sheet0 = workbook.getSheetAt(0);


            sheet.removeRow(sheet.getRow(0));
            sheet.removeRow(sheet.getRow(1));
            sheet.removeRow(sheet.getRow(2));

            sheet.createRow(0);
            sheet.createRow(1);
            sheet.createRow(2);*/

//            for (int i = 0; i < 2; i++) {
//                // Sao chép từng cell từ hàng nguồn sang hàng đích
//                Row srcRow = sheet0.getRow(i);
//                Row destRow = sheet.getRow(i);
//
//                for (int j = 0; j < srcRow.getLastCellNum(); j++) {
//                    Cell srcCell = srcRow.getCell(i);
//                    Cell destCell = destRow.createCell(i);
//
//                    if (srcCell != null) {
//                        destCell.setCellStyle(srcCell.getCellStyle());
//                        switch (srcCell.getCellType()) {
//                            case STRING -> destCell.setCellValue(srcCell.getStringCellValue());
//                            case NUMERIC -> destCell.setCellValue(srcCell.getNumericCellValue());
//                            case BOOLEAN -> destCell.setCellValue(srcCell.getBooleanCellValue());
//                            case FORMULA -> {
//                                String formula = srcCell.getCellFormula();
//                                destCell.setCellFormula(formula);
//                            }
//                            case BLANK -> destCell.setBlank();
//                            default -> {
//                            }
//                        }
//                    }
//                }
//            }

            // số cột chứa thông tin tính toán tự tạo sẽ ẩn đi khi đã nhập xong tính vật liệu để tránh rối
            int numColHide;
            // nếu số bozai < 15 thì số cột cần ẩn là 15, nếu không thì số cột ẩn là số bozai
            if (soBoZai < 15) {
                numColHide = 15;
            } else {
                numColHide = soBoZai;
            }
            // ẩn tất cả các cột từ numColHide + 8
            for (int i = numColHide + 8; i < sheet.getRow(6).getLastCellNum(); i++) {
                sheet.setColumnHidden(i, true);
            }

            // hợp nhất các ô
            // Xác định vùng cần hợp nhất (từ cột 6 đến cột 8 trên dòng 0)
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 6, 8);
            sheet.addMergedRegion(cellRangeAddress);

            cellRangeAddress = new CellRangeAddress(0, 0, 12, 14);
            sheet.addMergedRegion(cellRangeAddress);

            cellRangeAddress = new CellRangeAddress(1, 1, 2, 4);
            sheet.addMergedRegion(cellRangeAddress);

            cellRangeAddress = new CellRangeAddress(1, 1, 8, 10);
            sheet.addMergedRegion(cellRangeAddress);

            // Khóa sheet với mật khẩu
            sheet.protectSheet("");

            // Yêu cầu Excel tính toán lại tất cả các công thức khi tệp được mở
            ((XSSFWorkbook) workbook).setForceFormulaRecalculation(true);
            try (FileOutputStream fileOut = new FileOutputStream(excelCopyPath)) {
                workbook.write(fileOut);

                workbook.close();
            }


        } catch (IOException e) {
            if (e instanceof FileNotFoundException) {
                System.out.println("File đang được mở bởi người dùng khác");
                throw new FileNotFoundException();
            }
            System.out.println(e.getMessage());
            throw new RuntimeException(e);
        }


//        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
//        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        excelFileNames.add(new ExcelFile("Sheet " + sheetIndex + ": " + kouSyu, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    /**
     * chuyển các thông số của tính vật liệu trong file 3bc sang excel
     *
     * @param kaKouPairs     thông số chiều dài vật liệu và chiều dài + số lượng của các sản phẩm tính cho cây vật liệu đó
     * @param sheetIndex     thứ tự tạo sheet
     * @param excelFileNames tên file excel
     * @param workbook
     * @throws FileNotFoundException
     */
    private static void writeDataToExcelToriai2(Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs, int sheetIndex, ObservableList<ExcelFile> excelFileNames, Workbook workbook) throws FileNotFoundException {

        // tổng chiều dài các kozai
        double kouzaiChouGoukei = 0;
        double seiHinChouGoukei = 0;


        // nếu tên vật liệu có chứa [ thì phải đổi sang U vì tên này sẽ đặt tên cho sheet nên [ không dùng được
        if (kouSyu.contains("[")) {
            kouSyu = kouSyu.replace("[", "U");
        }

        // bản cũ copy excel từ file mẫu
            /*// Lấy index sheet gốc cần sao chép
            int sheetSampleIndex = 0;
            // sao chép sheet gốc sang một sheet mới
            workbook.cloneSheet(sheetSampleIndex);
            // đổi tên sheet mới theo tên vật liệu đang duyệt, sheetIndex là chỉ số của sheet mới
            workbook.setSheetName(sheetIndex, kouSyu);*/

        // lấy ra sheet đang duyệt cần chỉnh
        Sheet sheet = workbook.getSheetAt(sheetIndex - 1);

        // bản cũ copy excel từ file mẫu
            /*Date currentDate = new Date();
            SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");

            String time = sdf.format(currentDate);
            // Ghi thời gian hiện tại vào ô C1
            sheet.getRow(0).getCell(2).setCellValue(time);

            // Ghi tên khách hàng vào ô G6
            sheet.getRow(0).getCell(6).setCellValue(kyakuSakiMei);

            // Ghi bikou vào ô M12
            sheet.getRow(0).getCell(12).setCellValue(bikou);

            // Ghi shortNouKi vào ô S18
            sheet.getRow(0).getCell(18).setCellValue(shortNouKi);

            // Ghi saizu vào ô C2, chưa dùng
            sheet.getRow(1).getCell(2).setCellValue("");

            // Ghi chuyuBan vào ô I8
            sheet.getRow(1).getCell(8).setCellValue(chuyuBan);
            // Ghi teiHaiSha vào ô O14
            sheet.getRow(1).getCell(14).setCellValue(teiHaiSha);*/

        // lấy số loại bozai và sản phẩm
        int soBoZai = kaKouPairs.size();
        // tìm sheet có số lượng bozai lớn nhất
        soBozaiMax = Math.max(soBozaiMax, soBoZai);
        int soSanPham = seiHinList.size();

        // map chứa chiều dài và số lượng của các sản phẩm trong vật liệu đang tính của excel, chiều dài là key sẽ khong xuâ hiện lần 2
        // vì nếu có 2 chiều dài sẽ gộp làm 1 và cộng dồn số lượng để khớp với seiHinList của tính vật liệu 3bc, sau đó so sánh chiều dài số lượng
        // của nó với seiHinList và seiHinMap của 3bc, nếu khớp là tính vật liệu 3bc đúng và không báo lỗi
        // map này đã được xắp xếp vì chiều dài trong excel vốn đã được tự động động xắp xếp tăng dần
        Map<Double, Integer> seiHinMapExcel = new LinkedHashMap<>();


        // lấy hàng cuối cùng chứa dữ liệu trong cột a, chính là hàng cuối cùng chứa chiều dài số lượng sản phẩm
        int lastRowSeihin = getLastRowWithDataInColumn(sheet, COT_CHIEU_DAI_SAN_PHAM); // Cột A = index 0

        // lặp qua các hàng chứa chiều dài sản phẩm của excel và thêm vào map
        for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
            // lấy chiều dài sản phẩm
            Double seihinZenchou = Math.abs(Double.parseDouble(getStringNumberCellValue(sheet.getRow(i).getCell(COT_CHIEU_DAI_SAN_PHAM))));
            // lấy số lượng sản phẩm
            int seihinHonsuu = Math.abs((int) Double.parseDouble(getStringNumberCellValue(sheet.getRow(i).getCell(COT_SO_LUONG_SAN_PHAM))));

            // nếu sản phẩm đã có trong map thì lấy số lượng trong map rồi xóa sản phẩm đi rồi thêm lại sản phẩm với
            // số lượng trong map đã lấy + số lượng hiện tại
            // nếu chưa có trong map thì thêm sản phẩm với số lượng hiện tại
            if (seiHinMapExcel.get(seihinZenchou) != null) {
                int oldNum = seiHinMapExcel.get(seihinZenchou);
                seiHinMapExcel.remove(seihinZenchou);
                seiHinMapExcel.put(seihinZenchou, seihinHonsuu + oldNum);
            } else {
                seiHinMapExcel.put(seihinZenchou, seihinHonsuu);
            }
        }

        int sttSanPham = 0;
        // duyệt qua mảng chiều dài sản phẩm của excel và check xem có khớp với của 3bc không
        // nếu chỉ cần 1 cặp không khớp thì báo lỗi và dừng thực hiện
        for (Map.Entry<Double, Integer> sanPham : seiHinMapExcel.entrySet()) {
            double chieuDaiExcel = sanPham.getKey();
            // định dạng chiều dài excel về 1 chữ số phần thập phân và không làm tròn vì chiều dài của 3bc cũng dịnh dạng như vậy
            BigDecimal bdChieuDai3bc = new BigDecimal(chieuDaiExcel);
            chieuDaiExcel = bdChieuDai3bc.setScale(1, RoundingMode.DOWN).doubleValue(); // cắt (không làm tròn)

            double chieuDai3bc = seiHinList.get(sttSanPham);

            int soLuongExcel = sanPham.getValue();
            int soLuong3bc = seiHinMap.get(chieuDai3bc);

            if (chieuDaiExcel != chieuDai3bc || soLuongExcel != soLuong3bc) {

                throw new SecurityException();
            }
            sttSanPham++;
        }
        System.out.println("Các sản phẩm của vật liệu " + kouSyu + " trên Excel và 3bc khớp nhau");


        int doRongCuaCotTongChieuDaiSanPham = sheet.getColumnWidth(16);
        int doRongCuaCotSoSanPhamChuaTinh = sheet.getColumnWidth(17);

        // xóa và nhập lại kirirosu
        Cell b7 = sheet.getRow(7).getCell(1);
        b7.setBlank();
        b7.setCellValue(kirirosu);


        // xóa ở các ô bozai và số lượng của nó, xóa vùng nhập tính vật liệu
        for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM - 3; i <= lastRowSeihin; i++) {
            for (int j = 4; j <= 15; j++) {
                Row row = sheet.getRow(i);
                if (row == null) {
                    row = sheet.createRow(i);
                }
                Cell cell = row.getCell(j);
                if (cell == null) {
                    cell = row.createCell(j);
                }

                cell.setBlank();
            }
        }

        // cài đặt lại độ rộng các cột cho phù hợp
        int wiCol4 = sheet.getColumnWidth(4);
        int wiCol5 = sheet.getColumnWidth(5);
        int widthCol4Sum5Div2 = (wiCol4 + wiCol5) / 2;
        for (int i = 18; i < 33; i++) {
            sheet.setColumnWidth(i, widthCol4Sum5Div2);

        }
//        sheet.setColumnWidth(18, widthCol4Sum5Div2);
//        sheet.setColumnWidth(19, widthCol4Sum5Div2);
//        sheet.setColumnWidth(20, widthCol4Sum5Div2);
//        sheet.setColumnWidth(21, widthCol4Sum5Div2);
//        sheet.setColumnWidth(22, widthCol4Sum5Div2);
//        sheet.setColumnWidth(23, widthCol4Sum5Div2);
//        sheet.setColumnWidth(24, widthCol4Sum5Div2);

        // nếu số bozai nhiều hơn 6 bao nhiêu thì thêm số cột bozai với số lượng đó
        // copy và paste giá trị cho cột mới cho giống giá trị với các cột còn lại
//            soBoZai = 18;

        int soSanPhamCanThem = Math.max(soBoZai - 6, 0);
        if (soBoZai > 6) {

            // xóa sạch dữ liệu vùng chỉ định đang có dữ liệu cần ghi đè để tránh sau 1 cell trong vùng khi bị ghi đè mà lại gộp cell
            // sẽ gặp tình trạng báo lỗi nếu nhiều cell cùng có giá trị
            for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
                for (int j = 25; j <= 30; j++) {
                    sheet.getRow(i).getCell(j).setBlank();
                }
            }


            int soSanPhamExcel = lastRowSeihin - HANG_DAU_TIEN_CHUA_SAN_PHAM + 1;

            // lấy row sau hàng sản phẩm cuối
            Row sauHangSanPhamCuoi = sheet.getRow(lastRowSeihin + 1);
            if (sauHangSanPhamCuoi == null) {
                sauHangSanPhamCuoi = sheet.createRow(lastRowSeihin + 1);
            }

            // tạo cell sau hàng sản phẩm đầu tiên và ở cột vật liệu đầu tiên
            // dán công thức mẫu vào cell
            // các đoạn dưới tương tự
            Cell sauHangSanPhamCuoiVaCotVatLieuDauTien = sauHangSanPhamCuoi.getCell(4);
            if (sauHangSanPhamCuoiVaCotVatLieuDauTien == null) {
                sauHangSanPhamCuoiVaCotVatLieuDauTien = sauHangSanPhamCuoi.createCell(4);
            }
            sauHangSanPhamCuoiVaCotVatLieuDauTien.setCellFormula("SUM(Z9:AA" + (HANG_DAU_TIEN_CHUA_SAN_PHAM + soSanPhamExcel + 1) + ")");

            Row row7 = sheet.getRow(6);
            Row rowHangSanPhamDauTien = sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM);
            if (row7 == null) {
                row7 = sheet.createRow(6);
            }
            if (rowHangSanPhamDauTien == null) {
                rowHangSanPhamDauTien = sheet.createRow(HANG_DAU_TIEN_CHUA_SAN_PHAM);
            }

            Cell r7 = row7.getCell(17);
            if (r7 == null) {
                r7 = row7.createCell(17);
            }
            r7.setCellFormula("B7*SUM(AN7:BA7)/1000");

            Cell ao7 = row7.getCell(40);
            if (ao7 == null) {
                ao7 = row7.createCell(40);
            }
            // set align theo chiều ngang là trung tâm
            ao7.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            // gộp ô cột được dán với ô ở cột tiếp theo vì các công thức này nằm trên 2 ô cùng hàng
            CellRangeAddress cellRangeAddress = new CellRangeAddress(6, 6, 40, 40 + 1);
            sheet.addMergedRegion(cellRangeAddress);
            ao7.setCellFormula("E7*E8");


            Cell r10 = rowHangSanPhamDauTien.getCell(17);
            if (r10 == null) {
                r10 = rowHangSanPhamDauTien.createCell(17);
            }
            r10.setCellFormula("B10-SUM(AN10:BA10)");

            Cell z10 = rowHangSanPhamDauTien.getCell(25);
            if (z10 == null) {
                z10 = rowHangSanPhamDauTien.createCell(25);
            }
            z10.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            cellRangeAddress = new CellRangeAddress(HANG_DAU_TIEN_CHUA_SAN_PHAM, HANG_DAU_TIEN_CHUA_SAN_PHAM, 25, 25 + 1);
            sheet.addMergedRegion(cellRangeAddress);
            z10.setCellFormula("($A10+$B$8)*E10");

            Cell ao10 = rowHangSanPhamDauTien.getCell(40);
            if (ao10 == null) {
                ao10 = rowHangSanPhamDauTien.createCell(40);
            }
            ao10.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            cellRangeAddress = new CellRangeAddress(HANG_DAU_TIEN_CHUA_SAN_PHAM, HANG_DAU_TIEN_CHUA_SAN_PHAM, 40, 40 + 1);
            sheet.addMergedRegion(cellRangeAddress);
            ao10.setCellFormula("E10*E$8");

            Cell bd10 = rowHangSanPhamDauTien.getCell(55);
            if (bd10 == null) {
                bd10 = rowHangSanPhamDauTien.createCell(55);
            }
            bd10.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            bd10.setCellFormula("Q10");

            Cell be10 = rowHangSanPhamDauTien.getCell(56);
            if (be10 == null) {
                be10 = rowHangSanPhamDauTien.createCell(56);
            }
            be10.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
            be10.setCellFormula("R10");


            // sửa lại công thức cho 2 ô cuối cột Q và cột R theo công thức mới
            Cell oCuoiCotQ = sheet.getRow(lastRowSeihin + 1).getCell(16);
            if (oCuoiCotQ == null) {
                oCuoiCotQ = sauHangSanPhamCuoi.createCell(16);
            }
            oCuoiCotQ.setCellFormula("SUM(BD9:BD" + (lastRowSeihin + 2) + ")");

            Cell oCuoiCotR = sheet.getRow(lastRowSeihin + 1).getCell(17);
            if (oCuoiCotR == null) {
                oCuoiCotR = sauHangSanPhamCuoi.createCell(17);
            }
            oCuoiCotR.setCellFormula("SUM(BE9:BE" + (lastRowSeihin + 2) + ")");


            // từ các cell công thức mẫu gọi hàm dán công thức vào các vùng cần thêm, có vùng sẽ dựa theo số lượng sản phẩm
            // công thức thay đổi phù hợp theo ô và theo cột
            copySrcCellToRange(sheet, sauHangSanPhamCuoiVaCotVatLieuDauTien, 6, 15, 2, lastRowSeihin + 1, lastRowSeihin + 1, 1, true);
            copySrcCellToRange(sheet, ao7, 40, 51, 2, 6, 6, 1, true);
            copySrcCellToRange(sheet, r10, 17, 17, 1, HANG_DAU_TIEN_CHUA_SAN_PHAM, lastRowSeihin, 1, true);
            copySrcCellToRange(sheet, z10, 25, 36, 2, HANG_DAU_TIEN_CHUA_SAN_PHAM, lastRowSeihin, 1, true);
            copySrcCellToRange(sheet, ao10, 40, 51, 2, HANG_DAU_TIEN_CHUA_SAN_PHAM, lastRowSeihin, 1, true);
            copySrcCellToRange(sheet, bd10, 55, 55, 1, HANG_DAU_TIEN_CHUA_SAN_PHAM, lastRowSeihin, 1, true);
            copySrcCellToRange(sheet, be10, 56, 56, 1, HANG_DAU_TIEN_CHUA_SAN_PHAM, lastRowSeihin, 1, true);


            // code của bản excel mới nhưng hiện tại không dùng, đã thay bằng cách khác mã không cần sử dụng file excel mẫu
            // bằng cách tạo công thức tại các ô mẫu cố định đã biết trước rồi dùng hàm copySrcCellToRange,
            // hàm này có khả năng copy 1 cell và dán vào cả 1 vùng dữ liệu mà không cần quan tâm dán theo cột hay hàng, công thức vẫn thay đổi theo đúng chuẩn
            // dán toàn bộ các công thức ra vùng cần dán dựa theo số lượng sản phẩm đang có


/*
                InputStream sourceFile = ReadPDFToExcel.class.getResourceAsStream("/com/lenha/excel_3bc_toriai/sampleFiles/sample files2.xlsx");
                assert sourceFile != null;
                Workbook excelMau = new XSSFWorkbook(sourceFile);
                Sheet sheetMau = excelMau.getSheetAt(0);

                // xóa sạch dữ liệu vùng chỉ định đang có dữ liệu cần ghi đè để tránh sau 1 cell trong vùng khi bị ghi đè mà lại gộp cell
                // sẽ gặp tình trạng báo lỗi nếu nhiều cell cùng có giá trị
                for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
                    for (int j = 25; j <= 30; j++) {
                        sheet.getRow(i).getCell(j).setBlank();
                    }
                }
                // copy các hàng 7 và 10 từ cột 17 đến cột 58 của excel mẫu vào excel đang tạo
                // đây là các hàng có chứa công thức mẫu
                saoChepCacHangTrongVungCotTuFileMauVaoFileDich(sheetMau, sheet, new int[]{6, 9}, 17, 58, (XSSFWorkbook) workbook);

                // sửa lại công thức cho 2 ô cuối cột Q và cột R theo công thức mới
                Cell oCuoiCotQ = sheet.getRow(lastRowSeihin + 1).getCell(16);
                Cell oCuoiCotR = sheet.getRow(lastRowSeihin + 1).getCell(17);

                oCuoiCotQ.setCellFormula("SUM(BD9:BD" + (lastRowSeihin + 2) + ")");
                oCuoiCotR.setCellFormula("SUM(BE9:BE" + (lastRowSeihin + 2) + ")");

                // đến đây đã tạo được các hàng công thức mẫu của hàng sản phẩm đầu tiên, hàng 10, chỉ số hàng 9
                // chỉ có thể copy từ excel mẫu được hàng đầu tiên vì excel mẫu chỉ có 1 hàng, có 1 hàng không biết sheet cần chỉnh có bao nhiêu sản phẩm


                // tạo cell r10
                Cell r10 = sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getCell(17);

                // duyệt qua các hàng của các sản phẩm thứ 2 trở đi và thêm nốt công thức vào các hàng này
                // copy công thức từ hàng sản phẩm đầu tiên vào các hàng bên dưới lần lượt theo vòng lặp hàng
                for (int i = 0; i < soSanPhamExcel - 1; i++) {
                    if (r10 == null) continue;
                    // dán công thức theo cột R, chỉ số 17
                    // tạo cell của hàng sản phẩm đang lặp tại cột 17
                    Cell dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).getCell(17);
                    // gọi hàm copy, dán và update công thức theo hàng để dán công thức từ hàng đầu tiên vào,
                    // thêm tham số thứ 3 là khoảng cách hàng copy với hàng sản phẩm đầu tiên để biết cách thay đổi công thức cho phù hợp với hàng được dán
                    copyRowCellWithFormulaUpdate(r10, dstCell, i + 1);

                    Cell srcCell = null;
                    // copy các công thức từ cột AO đến cột AZ(cách nhau 12 cột) của hàng sản phẩm đầu tiên(HANG_DAU_TIEN_CHUA_SAN_PHAM) vào các hàng đang lặp
                    // copy các công thức từ cột Z đến cột AK(cách nhau 12 cột) của hàng sản phẩm đầu tiên(HANG_DAU_TIEN_CHUA_SAN_PHAM) vào các hàng đang lặp
                    for (int j = 0; j < 12; j++) {
                        // copy các công thức từ cột AO đến cột AZ(cách nhau 12 cột) của hàng sản phẩm đầu tiên(HANG_DAU_TIEN_CHUA_SAN_PHAM) vào các hàng đang lặp
                        // lấy cell của sản phẩm hàng đầu tiên tại cột đang lặp
                        srcCell = sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getCell(j + 40);// 40 là chỉ số cột AO
                        // lấy cell của sản phẩm cần dán tại hàng sản phẩm đang lặp và tại cột cần dán đang lặp
                        dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).getCell(j + 40);
                        // nếu cell chưa tồn tại thì tạo nó
                        if (dstCell == null) dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).createCell(j + 40);
                        // gọi hàm copy, dán và update công thức theo hàng để dán công thức từ hàng đầu tiên vào
                        copyRowCellWithFormulaUpdate(srcCell, dstCell, i + 1);
                        // gộp ô cột được dán với ô ở cột tiếp theo vì các công thức này nằm trên 2 ô cùng hàng
                        CellRangeAddress cellRangeAddress = new CellRangeAddress(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1, i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1
                                , j + 40, j + 40 + 1);
                        sheet.addMergedRegion(cellRangeAddress);


                        // tương tự như cách copy cột AO đến AZ ở trên
                        // copy các công thức từ cột Z đến cột AK(cách nhau 12 cột) của hàng sản phẩm đầu tiên(HANG_DAU_TIEN_CHUA_SAN_PHAM) vào các hàng đang lặp
                        // lấy cell của sản phẩm hàng đầu tiên tại cột đang lặp
                        srcCell = sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getCell(j + 25);// 25 là chỉ số cột Z
                        // lấy cell của sản phẩm cần dán tại hàng sản phẩm đang lặp và tại cột cần dán đang lặp
                        dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).getCell(j + 25);
                        // nếu cell chưa tồn tại thì tạo nó
                        if (dstCell == null) dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).createCell(j + 25);
                        // gọi hàm copy, dán và update công thức theo hàng để dán công thức từ hàng đầu tiên vào
                        copyRowCellWithFormulaUpdate(srcCell, dstCell, i + 1);
                        // gộp ô cột được dán với ô ở cột tiếp theo vì các công thức này nằm trên 2 ô cùng hàng
                        cellRangeAddress = new CellRangeAddress(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1, i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1
                                , j + 25, j + 25 + 1);
                        sheet.addMergedRegion(cellRangeAddress);

                        // tự động tăng j để bỏ qua 1 cột vì ô ở cột tiếp theo đã được gộp với ô ở cột hiện tại
                        j++;
                    }

                    // copy các công thức từ cột BD đến cột BE(cách nhau 2 cột) của hàng sản phẩm đầu tiên(HANG_DAU_TIEN_CHUA_SAN_PHAM) vào các hàng đang lặp
                    for (int j = 0; j < 2; j++) {
                        // lấy cell của sản phẩm hàng đầu tiên tại cột đang lặp
                        srcCell = sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getCell(j + 55);// 55 là chỉ số cột BD
                        // lấy cell của sản phẩm cần dán tại hàng sản phẩm đang lặp và tại cột cần dán đang lặp
                        dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).getCell(j + 55);
                        // nếu cell chưa tồn tại thì tạo nó
                        if (dstCell == null) dstCell = sheet.getRow(i + HANG_DAU_TIEN_CHUA_SAN_PHAM + 1).createCell(j + 55);
                        // gọi hàm copy, dán và update công thức theo hàng để dán công thức từ hàng đầu tiên vào
                        copyRowCellWithFormulaUpdate(srcCell, dstCell, i + 1);
                    }

                }
*/
            // xóa hợp nhất các ô vùng thông tin để không bị lỗi khi thêm cột, sau khi làm hết việc sẽ fomat lại hợp nhất ô như ban đầu
            unmergeAndFillCellsInRange(sheet, 0, 2, 0, 17);



            // lưu giá trị công thức cũ của ô công thức gốc này mỗi lần lặp để gán lại cho nó khi công thức của nó bị thay đổi khôn đúng
            String congThucSauHangSanPhamCuoiVaCotVatLieuDauTien = "SUM(Z9:AA" + (HANG_DAU_TIEN_CHUA_SAN_PHAM + soSanPhamExcel + 1) + ")";
            // thêm j lần các cột mới tại các cột công thức
            int soBozaiCanThem = soBoZai - 6;
            for (int j = 0; j < soBozaiCanThem; j++) {
                sheet.shiftColumns(42 + 4 * j, sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getLastCellNum(), 2);
                sheet.shiftColumns(27 + 2 * j, sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getLastCellNum(), 2);
                // sau khi dịch chuyển các cột ở dòng code trên thì công thức của sauHangSanPhamCuoiVaCotVatLieuDauTien đã bị thay đổi
                // nên cần gán lại công thức cũ của nó cho đúng
                sauHangSanPhamCuoiVaCotVatLieuDauTien.setCellFormula(congThucSauHangSanPhamCuoiVaCotVatLieuDauTien);
                sheet.shiftColumns(6, sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM).getLastCellNum(), 2);
//                System.out.println("last col: " + sheet.getRow(6).getLastCellNum());

                // lấy lại công thức của cell sauHangSanPhamCuoiVaCotVatLieuDauTien trong vòng lặp này để
                // khi nó bị thay đổi công thức sai trong vòng lặp sau thì gán lại công thức đúng này cho nó
                congThucSauHangSanPhamCuoiVaCotVatLieuDauTien = sauHangSanPhamCuoiVaCotVatLieuDauTien.getCellFormula();

                // dịch chuyển các cột ở 3 hàng tiêu đề và hàng tên người làm về vị trí ban đầu sau khi bị dịch chuyển sang phải 2 hàng
                // nếu cell đã tồn tại thì cũng không thể thực hiện phép lùi cột, khi dịch cột mà cell ở vùng tạo thêm chưa được tạo thì mới lùi cột được
                for (int i = 0; i < sheet.getLastRowNum(); i++) {
                    if (i < 3 || i > lastRowSeihin + 5) {
                        Row row = sheet.getRow(i);
                        row.shiftCellsLeft(48 + j * 4, 10000, 2);
                        row.shiftCellsLeft(31 + j * 2, 10000, 2);

                        row.shiftCellsLeft(8, 10000, 2);

                    }

                }

//
//                    // sửa lại công thức tất cả các ô có giá trị L về K vì sau khi dịch chuyển 3 hàng tiêu đề về vị trí ban đầu
//                    // công thức bị sai
//                    for (int i = 26 + j; i <= 41 + 2 * j; i++) {
//                        // row index 6 tức là hàng 7 vì ban đầu chỉ có 1 hàng 7 có công thức do chưa thêm các hàng mới
//                        Row row = sheet.getRow(6);
//                        Cell cell = row.getCell(i);
//
//                        if (cell != null && cell.getCellType() == CellType.FORMULA) {
//                            String formula = cell.getCellFormula();
//                            formula = formula.replaceAll("\\$L\\$3", "\\$K\\$3");
//                            cell.setCellFormula(formula);
//                        }
//                    }
//
                Cell srcCell;
                Cell destCell;

                // lặp qua các hàng cần sao chép
                // sao chép ô từ cột 5, 6 sang cột 7, 8 từ hàng 4 đến hàng cuối chứa sản phẩm
                // cần tạo cell ở cột 7, 8 bị phép dịch chuyển cột ở trên thực chất chưa tạo cell mới
                // nếu cell đã tồn tại thì cũng không thể thực hiện phép lùi cột, khi dịch cột mà cell ở vùng tạo thêm chưa được tạo thì mới lùi cột được
                for (int i = 3; i <= lastRowSeihin + 5; i++) {
                    Row row = sheet.getRow(i);
                    // Sao chép ô từ cột srcColumn sang destColumn trong các cột tính vật liệu sản phẩm
                    srcCell = row.getCell(4);
                    if (srcCell == null) {
                        srcCell = row.createCell(4);
                    }
                    int destCol = 6;
                    row.createCell(destCol);
                    row.createCell(destCol + 1);

                    copySrcCellToRange(sheet, srcCell, destCol, destCol, 1, i, i, 1, true);
                    srcCell = row.getCell(5);
                    if (srcCell == null) {
                        srcCell = row.createCell(5);
                    }
                    destCell = row.getCell(7);
                    destCell.setCellStyle(srcCell.getCellStyle());

                    // Sao chép ô từ cột srcColumn sang destColumn trong các cột công thức tự tạo đầu tiên
                    // vì thực hiện sau khi đã thêm 2 cột ở sản phẩm làm vùng này tăng lên 2 cột nên ô công thức gốc
                    // có chỉ số cột không còn là 25 nữa mà là 27, còn ô cần dán không còn 27 nữa mà là 29,
                    // rồi do tính theo số lần thêm cột ở cột sản phẩm nên thêm j * 2
                    srcCell = row.getCell(27 + j * 2);
                    if (srcCell == null) {
                        srcCell = row.createCell(27 + j * 2);
                    }
                    destCol = 29 + j * 2;
                    row.createCell(destCol);
                    row.createCell(destCol + 1);
                    copySrcCellToRange(sheet, srcCell, destCol, destCol, 1, i, i, 1, true);

                    // Sao chép ô từ cột srcColumn sang destColumn trong các cột công thức tự tạo thứ hai vì thực hiện
                    // sau khi đã thêm 2 cột ở sản phẩm và 2 cột ở công thức tự tạo đầu tiên làm vùng này tăng lên 2 + 2 cột
                    // nên ô công thức gốc có chỉ số cột không còn là 40 nữa mà là 40 + 2 + 2 = 44, còn ô cần dán
                    // không còn 42 nữa mà là 42 + 2 + 2 = 46, rồi do tính theo số lần thêm 2 cột ở sản phẩm và 2 cột ở công thức tự tạo đầu tiên nên thêm j * (2 + 2)
                    srcCell = row.getCell(44 + j * 4);
                    if (srcCell == null) {
                        srcCell = row.createCell(44 + j * 4);
                    }
                    destCol = 46 + j * 4;
                    row.createCell(destCol);
                    row.createCell(destCol + 1);
                    copySrcCellToRange(sheet, srcCell, destCol, destCol, 1, i, i, 1, true);

                }

//                    // tại hàng 7 copy ô từ cột 26 sang 27
//                    Row row7Formula = sheet.getRow(6);
//                    srcCell = row7Formula.getCell(26 + j);
//                    destCell = row7Formula.createCell(27 + j);
//                    copyCellWithFormulaUpdate(srcCell, destCell, 1);
//
//                    // tại hàng 7 copy ô từ cột 44 sang 45
//                    srcCell = row7Formula.getCell(44 + 2 * j);
//                    destCell = row7Formula.createCell(45 + 2 * j);
//                    copyCellWithFormulaUpdate(srcCell, destCell, 1);
//
//                    // tại hàng 4 copy ô từ cột 44 sang 45
//                    Row row4Formula = sheet.getRow(3);
//                    srcCell = row4Formula.getCell(44 + 2 * j);
//                    destCell = row4Formula.createCell(45 + 2 * j);
//                    copyCellWithFormulaUpdate(srcCell, destCell, 1);
            }

            try {
                // hợp nhất các ô
                // sau khi thêm các cột vật liệu sẽ làm các ô thông tin mất hợp nhất nên cần hợp nhất lại
                // chỉ thực hiện trong sheet có số bozai > 6 vì chỉ nó mới thêm cột, còn các sheet khác không thêm nên không mất hợp nhất, nếu tiếp tục thực lệnh sẽ không thể ghi đè và gây ra lỗi
                // Xác định vùng cần hợp nhất (từ cột 6 đến cột 8 trên dòng 0)
                cellRangeAddress = new CellRangeAddress(0, 0, 0, 1);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(0, 0, 2, 4);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(0, 0, 8, 10);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(0, 0, 14, 16);
                sheet.addMergedRegion(cellRangeAddress);

                cellRangeAddress = new CellRangeAddress(1, 1, 0, 1);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(1, 1, 2, 4);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(1, 1, 8, 10);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(1, 1, 14, 16);
                sheet.addMergedRegion(cellRangeAddress);

                cellRangeAddress = new CellRangeAddress(2, 2, 0, 1);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(2, 2, 2, 4);
                sheet.addMergedRegion(cellRangeAddress);
                cellRangeAddress = new CellRangeAddress(2, 2, 8, 16);
                sheet.addMergedRegion(cellRangeAddress);
            } catch (Exception e) {
                System.out.println("Lỗi hợp nhất ô");
            }


        }


        // cài đặt lại độ rộng các cột vật liệu mới được thêm vào cho giống với độ rộng của các cột vật liệu cũ
        int widthCol5 = sheet.getColumnWidth(4);
        int widthCol6 = sheet.getColumnWidth(5);
        for (int i = 6; i <= 15 + soSanPhamCanThem * 2; i += 2) {
            sheet.setColumnWidth(i, widthCol5);
            sheet.setColumnWidth(i + 1, widthCol6);
        }

        // cài đặt lại độ rộng các cột tổng chiều dài sản phẩm và số lượng của sản phẩm chưa tính vật liệu vì sau khi
        // thêm các cột sản phẩm mới nó bị dịch chuyển sang cột khác
        sheet.setColumnWidth(16 + soSanPhamCanThem * 2, doRongCuaCotTongChieuDaiSanPham);
        sheet.setColumnWidth(17 + soSanPhamCanThem * 2, doRongCuaCotSoSanPhamChuaTinh);

        int witdthCol23 = sheet.getColumnWidth(23);
        // Bỏ ẩn các cột công thức để khi thêm cột bozai nếu trùng với các cột công thức cũ thì nó không bị ẩn
        // Bỏ ẩn một dải cột (ví dụ cột E..G -> index 4..6)
        for (int col = 24; col <= 31; col++) {
            sheet.setColumnHidden(col, false);

            // Nếu cột trước đó bị set width = 0, thì đặt lại width hợp lý
            if (sheet.getColumnWidth(col) == 0) {
                sheet.setColumnWidth(col, witdthCol23); // 15 ký tự ~ 15*256
            }
        }


        // xóa giá trị tất cả các hàng sản phẩm trong excel
        for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
            sheet.getRow(i).getCell(0).setBlank();
            sheet.getRow(i).getCell(1).setBlank();
        }

        // ghi tất cả sản phẩm vào excel
        int indexSeiHin = HANG_DAU_TIEN_CHUA_SAN_PHAM;
        for (Map.Entry<Double, Integer> sanPham : seiHinMapExcel.entrySet()) {
            sheet.getRow(indexSeiHin).getCell(0).setCellValue(sanPham.getKey());
            sheet.getRow(indexSeiHin).getCell(1).setCellValue(sanPham.getValue());
            indexSeiHin++;
        }


            /*// nếu số sản phẩm lớn hơn 1 bao nhiêu lần thì thêm số hàng sản phẩm số lần tương tự
            if (soSanPham > 1) {
                for (int j = 0; j < soSanPham - 1; j++) {
                    // đẩy tất cả các hàng ở dưới hàng index 6 xuống 1 hàng để thừa ra hàng index 7 nhưng nó thực tế vẫn chưa được tạo
                    // sau đó mới tạo hàng index 7
                    sheet.shiftRows(7, sheet.getLastRowNum(), 1);
                    Row srcRow = sheet.getRow(6);
                    // tạo hàng index 7
                    Row destRow = sheet.createRow(7);

                    // Sao chép từng cell từ hàng nguồn sang hàng đích
                    for (int i = 0; i < srcRow.getLastCellNum(); i++) {
                        Cell sourceCell = srcRow.getCell(i);
                        Cell targetCell = destRow.createCell(i);

                        if (sourceCell != null) {
                            copyRowCellWithFormulaUpdate(sourceCell, targetCell, 1);
                        }
                    }
                }
            }*/

            /*// ghi tất cả sản phẩm vào excel, cách cũ với các chiều dài sản phẩm là duy nhất do ta tự thực hiện ghi, còn cách mới là dùng chiều dài đã ghi sẵn trên excel nên không thực hiện được
            for (int i = 0; i < soSanPham; i++) {
                Double length = seiHinList.get(i);
                // ghi chiều dài sản phẩm
                sheet.getRow(i + 6).getCell(0).setCellValue(length);
                // ghi số lượng sản phẩm
                sheet.getRow(i + 6).getCell(1).setCellValue(seiHinMap.get(length));
            }*/


        // ghi bozai và sản phẩm trong bozai
        // thể hiện index cột bozai đang thực thi
        int numBozai = 0;
        // lặp qua các cặp tính vật liệu, mỗi cặp gồm "bozai-số lượng trong map kouZaiChouPairs(nằm trong map nhưng thực tế nó chỉ có 1 cặp)"
        // và "các bộ chiều dài sản phẩm và số lượng trong map meiSyouPairs"
        for (Map.Entry<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> entry : kaKouPairs.entrySet()) {
            // lấy chỉ số cột vật liệu đang lặp
            int colBozai = 4 + numBozai;

            Map<StringBuilder, Integer> kouZaiChouPairs = entry.getKey();
            Map<StringBuilder[], Integer> meiSyouPairs = entry.getValue();

            // chiều dài vật liệu đang xét
//                double chieuDaiBozai;
            // số lượng của vật liệu đang xét
            int soLuongCuaBozai = 0;

            // Ghi bozai và số lượng của nó, thực ra chỉ có 1 cặp bozai và số lượng nhưng vẫn tạo vòng lặp cho đơn giản, không ảnh hưởng gì
            for (Map.Entry<StringBuilder, Integer> kouZaiEntry : kouZaiChouPairs.entrySet()) {
                String chieuDaiBoZaiS = String.valueOf(kouZaiEntry.getKey());
                String soLuongCuaBozaiS = String.valueOf(kouZaiEntry.getValue());

//                    chieuDaiBozai = Double.parseDouble(String.valueOf(kouZaiEntry.getKey()));
                soLuongCuaBozai = Integer.parseInt(soLuongCuaBozaiS);

                sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM - 3).getCell(colBozai, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(chieuDaiBoZaiS);
                sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM - 2).getCell(colBozai, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(soLuongCuaBozaiS);

                kouzaiChouGoukei += Double.parseDouble(chieuDaiBoZaiS) * Integer.parseInt(soLuongCuaBozaiS);
            }

            // lặp qua các chiều dài sản phẩm trong cặp tính vật liệu này và tìm trong hàng có chiều dài sản phẩm
            // tương ứng trong cột sản phẩm, dóng sang cột bozai đang tạo là tìm được ô cần ghi số lượng, sau đó ghi cộng dồn số lượng vào ô đó
            for (Map.Entry<StringBuilder[], Integer> meiSyouEntry : meiSyouPairs.entrySet()) {
                // chiều dài sản phẩm trong map
                double chieuDaiSanPhamTrongMap = Double.parseDouble(meiSyouEntry.getKey()[1].toString());
                // số lượng sản phẩm trong map
                int soLuongSanPhamTrongMap = Integer.parseInt(meiSyouEntry.getValue().toString());

                // cách mới khi một chiều dài chỉ xuất hiện 1 lần trên 1 hàng, khác với cách cũ phức tạp hơn
                // hàng chứa sản phẩm, +9 vì cột chứa sản phẩm bắt đầu chứa các sản phẩm từ hàng thứ 9
                int indexSeiHinRow = seiHinList.indexOf(chieuDaiSanPhamTrongMap) + HANG_DAU_TIEN_CHUA_SAN_PHAM;

                // lấy cell chứa số lượng của sản phẩm ở cột bozai
                Cell cellSoLuong = sheet.getRow(indexSeiHinRow).getCell(4 + numBozai);

                // lấy số lượng cũ của cell
                double oldNum = 0d;
                // nếu cell có type là số thì nó đã có số lượng từ trước thì gán nó cho số lượng cũ
                if (cellSoLuong.getCellType() == CellType.NUMERIC) {
                    oldNum = cellSoLuong.getNumericCellValue();
                }

                // nếu số lượng cũ > 0 thì ghi giá trị cell với số lượng cũ + số lượng hiện tại
                // không thì ghi cell với số lượng hiện tại
                if (oldNum > 0d) {
                    sheet.getRow(indexSeiHinRow).getCell(4 + numBozai).setCellValue(soLuongSanPhamTrongMap + oldNum);
                } else {
                    sheet.getRow(indexSeiHinRow).getCell(4 + numBozai).setCellValue(soLuongSanPhamTrongMap);
                }
                // kết thúc cách mới


// cách cũ khi nhiều hàng sản phẩm có chiều dài giống nhau, phức tạp hơn bình thường, vẫn hoạt động nhưng với trường hợp gộp các hàng chiều dài giống nhau thì không cần thiết vì làm thời gian chạy lâu hơn
//
//                // lặp qua các hàng chiều dài sản phẩm và so sánh với chiều dài sản phẩm trong map, nếu khớp nhau thì tính toán tại hàng đó, thêm số lượng
//                // sản phẩm tương ứng vào ô số lượng tại hàng tìm thấy trên cột bozai đang lặp
//
//                // biến nhớ số lượng còn lại của sản phẩm trong map khi lặp qua các hàng sản phẩm để nhập số lượng của sản phẩm trong map này
//                int soLuongConLaiCuaSanPhamTrongMap = soLuongSanPhamTrongMap;
//
//                // cách để ghi số lượng các sản phẩm khác nhau nhưng chiều dài giống nhau không ghi số lượng dồn vào chỉ 1 sản phẩm khi tính
//                // nhưng cách này chưa hoạt động
//                // có thể debug để xem lại, có thể tính tới cách khác là tạo tên sản phẩm bằng số thứ tự trong 3bc rồi dùng tên đó để so sánh giống thì nhập số lượng(chưa thử)
////                    Map<Double, Integer> hangDaGHiLanTruoc = new HashMap<>();
//
//                for (int i = HANG_DAU_TIEN_CHUA_SAN_PHAM; i <= lastRowSeihin; i++) {
//
//
//                    // lấy chiều dài sản phẩm trong hàng đang lặp tại cột sản phẩm
//                    double chieuDaisanPham = Math.abs(Double.parseDouble(getStringNumberCellValue(sheet.getRow(i).getCell(COT_CHIEU_DAI_SAN_PHAM))));
//
//                    // cách để ghi số lượng các sản phẩm khác nhau nhưng chiều dài giống nhau không ghi số lượng dồn vào chỉ 1 sản phẩm khi tính
//                    // nhưng cách này chưa hoạt động
//                        /*Integer chiSoHangDaGHiLanTruoc;
//                        chiSoHangDaGHiLanTruoc = hangDaGHiLanTruoc.get(chieuDaisanPham);
//                        if (chiSoHangDaGHiLanTruoc != null && i == chiSoHangDaGHiLanTruoc) {
//                            continue;
//                        }*/
//
//                    // nếu chiều dài sản phẩm tại hàng đang lặp khớp với chiều dài sản phẩm trong map thì bắt đầu tính toán
//                    if (chieuDaisanPham == chieuDaiSanPhamTrongMap) {
//
//                        // lấy số lượng sản phẩm tại cột sản phẩm
//                        int soLuongSanPham = Math.abs((int) Double.parseDouble(getStringNumberCellValue(sheet.getRow(i).getCell(COT_SO_LUONG_SAN_PHAM))));
//
//                        // đoạn code này tính toán số lượng còn lại của sản phẩm
//                        // biến lưu số lượng đã tính của sản phẩm này trước khi tính tiếp ở bước dưới
//                        int soLuongDatinhCuaSanPham = 0;
//                        // lặp qua toàn bộ các cột bozai tại hàng sản phẩm này
//                        // lấy tổng số lượng đã tính vật liệu của sản phẩm này
//                        // bằng cách tại cột đang lặp nhân số lượng của bozai với số lượng của sản phẩm rồi cộng tổng kết quả các lần lặp là ra tổng số lượng sản phẩm đã tính vật liệu
//                        for (int j = 4; j <= 3 + soBoZai * 2; j += 2) {
//                            // lấy kết quả dạng chuỗi của số lượng bozai và số lượng sản phẩm tại cột bozai đang lặp
//                            String slBZ = getStringNumberCellValue(sheet.getRow(HANG_DAU_TIEN_CHUA_SAN_PHAM - 2).getCell(j));
//                            String slSP = getStringNumberCellValue(sheet.getRow(i).getCell(j));
//
//                            // gán lại giá trị số lượng tương ứng dạng int
//                            // số lượng tại hàng bozai trong cột bozai đang lặp
//                            int soLuongTaiHangBozai = 0;
//                            if (!slBZ.equalsIgnoreCase("")) {
//                                soLuongTaiHangBozai = Math.abs((int) Double.parseDouble(slBZ));
//                            }
//
//                            // số lượng tại hàng sản phẩm trong cột bozai đang lặp
//                            int soLuongTaiHangSanPham = 0;
//                            if (!slSP.equalsIgnoreCase("")) {
//                                soLuongTaiHangSanPham = Math.abs((int) Double.parseDouble(slSP));
//                            }
//
//                            // cộng dồn số lượng tính tại cột bozai này vào số lượng đã tính của sản phẩm
//                            soLuongDatinhCuaSanPham += soLuongTaiHangBozai * soLuongTaiHangSanPham;
//                        }
//
//                        // tính số lượng còn lại của sản phẩm bằng số lượng ban đầu - số lượng đã tính
//                        int soLuongConLaiCuaSanPham = soLuongSanPham - soLuongDatinhCuaSanPham;
//
//                        // cách khác lấy giá trị số lượng còn lại của sản phẩm trong hàng sản phẩm này bằng giá trị ở cột hiển thị
//                        // tuy nhiên nó không hoạt động vì công thức không hề cập nhật giá trị khi ghi số lượng vào các ô tính vật liệu làm công thức bị sai và chương trình chạy không đúng nữa
//                       /*     int soLuongConLaiCuaSanPham = 0;
//                            Cell cellSlConLaiCuaSP = sheet.getRow(i).getCell(15 + 2 + (soBoZai - 6) * 2);
//                            String slConLaiCuaSP = getStringNumberCellValue(cellSlConLaiCuaSP);
//                            if (!slConLaiCuaSP.equalsIgnoreCase("")){
//                                soLuongConLaiCuaSanPham = (int) Double.parseDouble(slConLaiCuaSP);
//                            }*/
//
//                        // nếu số lượng còn lại không còn thì bỏ qua hàng sản phẩm này chuyển xuống hàng sản phẩm tiếp phía dưới
//                        if (soLuongConLaiCuaSanPham > 0) {
//                            // lấy cell chứa số lượng cũ của sản phẩm trong hàng sản phẩm tại trong cột bozai đang lặp
//                            Cell cellSoLuong = sheet.getRow(i).getCell(colBozai);
//                            // lấy số lượng cũ của cell trong hàng sản phẩm và trong cột bozai đang lặp
//                            int soLuongCuCuaSanPhamTaiCotBozai = 0;
//                            // nếu cell có type là số thì nó đã có số lượng từ trước thì gán nó cho số lượng cũ
//                            if (cellSoLuong.getCellType() == CellType.NUMERIC) {
//                                soLuongCuCuaSanPhamTaiCotBozai = (int) cellSoLuong.getNumericCellValue();
//                            }
//
//                            // tính số lượng sản phẩm có thể ghi bằng số lượng còn lại của sản phẩm / số lượng của bozai tại cột bozai đang lặp
//                            int soLuongCoTheGhi = soLuongConLaiCuaSanPham / soLuongCuaBozai;
//
//                            // nếu số lượng < 1 tức không thể ghi nữa thì chuyển xuống hàng sản phẩm tiếp theo
//                            if (soLuongCoTheGhi < 1) {
//                                continue;
//                            }
//
//                            //  nếu số lượng có thể ghi lớn hơn số lượng còn lại của sản phẩm trong map thì
//                            // số lượng có thể ghi gán lại nhỏ hơn = số lượng còn lại của sản phẩm trong map
//                            if (soLuongCoTheGhi > soLuongConLaiCuaSanPhamTrongMap) {
//                                soLuongCoTheGhi = soLuongConLaiCuaSanPhamTrongMap;
//                            }
//
//                            // sau khi tính xong số lượng có thể ghi thì trừ số lượng còn lại của sản phẩm trong map cho số lượng có thể ghi
//                            // vì chắc chắn nó sẽ ghi thêm vào số lượng nên cần trừ đi
//                            soLuongConLaiCuaSanPhamTrongMap -= soLuongCoTheGhi;
//
//
//                            // đoạn code không chắc có chạy không nhưng cũng không cần thiết
//                                /*int soLuongConLaiCuaSanPhamTrongMap;
//                                if (soLuongConLaiCuaSanPham <= soLuongSanPhamTrongMap) {
//                                    soLuongConLaiCuaSanPhamTrongMap = soLuongConLaiCuaSanPham;
//                                } else {
//                                    soLuongConLaiCuaSanPhamTrongMap = soLuongSanPhamTrongMap;
//                                }
//                                // số lượng cần ghi sẽ phải là nó chia cho số lượng của bozai tại cột đang lặp
//                                // để tính xem hàng sản phẩm này có đúng là hàng cần ghi không vì có thể còn hàng sản phẩm khác có chiều dài tương tự cần ghi
//                                soLuongConLaiCuaSanPhamTrongMap = soLuongConLaiCuaSanPhamTrongMap / soLuongCuaBozai;
//
//                                // tính lại số lượng còn lại của sản phẩm trong map để dùng cho các hàng tiếp theo nếu còn chiều dài giống hàng này
//                                soLuongSanPhamTrongMap -= soLuongConLaiCuaSanPhamTrongMap * soLuongCuaBozai;*/
//
//                            // đến đây mà vẫn chưa bị thoát thì mọi điều kiện ph hợp, ghi nối số lượng sản phẩm trong map vào ô của tính vật liệu của sản phẩm trong hàng đang lặp
//                            // sau đó thoát khỏi vòng lặp này để nhảy sang vòng bên ngoài và duyệt tiếp chiều dài sản phẩm khác trong map của bozai đang duyệt của vòng lặp ngoài
//                            // nếu số lượng cũ > 0 thì ghi giá trị cell với số lượng cũ + số lượng hiện tại
//                            // không thì ghi cell với số lượng hiện tại
//                            if (soLuongCuCuaSanPhamTaiCotBozai > 0d) {
//                                sheet.getRow(i).getCell(4 + numBozai).setCellValue(soLuongCoTheGhi + soLuongCuCuaSanPhamTaiCotBozai);
//                            } else {
//                                sheet.getRow(i).getCell(4 + numBozai).setCellValue(soLuongCoTheGhi);
//                            }
//
//                            // cách để ghi số lượng các sản phẩm khác nhau nhưng chiều dài giống nhau không ghi số lượng dồn vào chỉ 1 sản phẩm khi tính
//                            // nhưng cách này chưa hoạt động
////                                hangDaGHiLanTruoc.put(chieuDaisanPham, i);
//
//                            // nếu số lượng còn lại của sản phẩm trong map bị trừ về 0 thì đã nhập xong sản phẩm này nên thoát
//                            if (soLuongConLaiCuaSanPhamTrongMap <= 0) {
//                                break;
//                            }
//
//
//                        }
//                    }
//
//                }


                // thống kê phục vụ cho hiển thị thông tin trên phầm mềm
//                    double totalLength = chieuDaiSanPhamTrongMap * soLuongSanPhamTrongMap;
//                    Cell cellSoLuongBozai = sheet.getRow(4).getCell(3 + numBozai);
//                    if (cellSoLuongBozai.getCellType() == CellType.STRING) {
//                        totalLength *= Double.parseDouble(cellSoLuongBozai.getStringCellValue());
//                    }
// kết thúc cách cũ
                seiHinChouGoukei += chieuDaiSanPhamTrongMap * soLuongSanPhamTrongMap * soLuongCuaBozai;

            }

            numBozai += 2;
        }

        // số cột chứa thông tin tính toán tự tạo sẽ ẩn đi khi đã nhập xong tính vật liệu để tránh rối
        int soCotBozai;
        // nếu số bozai < 6 thì số cột cần ẩn là từ cột 6 * 2 + 12 đến hết, nếu không thì số cột ẩn là từ cột số bozai * 2 + 12 đến hết
        if (soBoZai < 6) {
            soCotBozai = 6 * 2;
        } else {
            soCotBozai = soBoZai * 2;
        }
        int lastCol = sheet.getRow(9).getLastCellNum();
        maxLastColNum = Math.max(maxLastColNum, lastCol);
        // ẩn tất cả các cột từ soCotBozai + 8, tạm thời không dùng vì sau khi xắt các sheet khác vào sheet 0 thì nó sẽ làm ẩn các cột của sheet 0
        // nếu sheet 0 có nhiều cột vật liệu hơn các sheet khác
/*        for (int i = soCotBozai + 12; i <= lastCol; i++) {
            sheet.setColumnHidden(i, true);
        }*/

        makeTextWhite(sheet, soCotBozai + 12, lastCol, -1, -1);


            /*// Khóa sheet với mật khẩu
            sheet.protectSheet("");*/


//        System.out.println("tong chieu dai bozai " + kouzaiChouGoukei);
//        System.out.println("tong chieu dai san pham " + seiHinChouGoukei);
        excelFileNames.add(new ExcelFile("Sheet " + sheetIndex + ": " + kouSyu, kouSyuName, kouzaiChouGoukei, seiHinChouGoukei));

    }

    /**
     * Đặt màu chữ thành trắng trong phạm vi cột/hàng cho trước.
     * @param sheet     Sheet cần sửa (Apache POI Sheet)
     * @param startCol  index cột bắt đầu (zero-based)
     * @param endCol    index cột kết thúc (zero-based)
     * @param startRow  index hàng bắt đầu (zero-based). Nếu = -1 => 0
     * @param endRow    index hàng kết thúc (zero-based). Nếu = -1 => sheet.getLastRowNum()
     */
     private static void makeTextWhite(Sheet sheet, int startCol, int endCol, int startRow, int endRow) {
        if (sheet == null) return;

        // chuẩn hóa cột (hoán vị nếu nhập ngược)
        if (startCol > endCol) {
            int t = startCol; startCol = endCol; endCol = t;
        }

        // xử lý row = -1
        if (startRow == -1) startRow = 0;
        if (endRow == -1) endRow = sheet.getLastRowNum();

        // nếu không hợp lệ thì thoát
        if (startRow > endRow || startCol < 0 || endCol < 0) return;

        Workbook wb = sheet.getWorkbook();

        // tạo 1 CellStyle + Font màu trắng dùng chung cho toàn vùng (tái sử dụng để tránh tạo quá nhiều style)
        CellStyle whiteStyle = wb.createCellStyle();
        Font whiteFont = wb.createFont();
        whiteFont.setColor(IndexedColors.WHITE.getIndex());
        whiteStyle.setFont(whiteFont);

        // duyệt hàng và cột — chỉ áp style cho ô đã tồn tại
        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = startCol; c <= endCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) continue;
                cell.setCellStyle(whiteStyle);
            }
        }
    }

    public static void writeWithoutCalcChain(Workbook wb, File outFile) throws IOException {
        // 1) write workbook to byte[] in memory
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos);
        wb.close();
        byte[] xlsxBytes = bos.toByteArray();

        // 2) open as zip, copy entries except xl/calcChain.xml
        try (ZipInputStream zin = new ZipInputStream(new ByteArrayInputStream(xlsxBytes));
             ZipOutputStream zout = new ZipOutputStream(new FileOutputStream(outFile))) {

            ZipEntry entry;
            byte[] buffer = new byte[8192];
            while ((entry = zin.getNextEntry()) != null) {
                String name = entry.getName();
                if ("xl/calcChain.xml".equals(name)) {
                    // skip it
                    zin.closeEntry();
                    continue;
                }
                // copy entry
                zout.putNextEntry(new ZipEntry(name));
                int len;
                while ((len = zin.read(buffer)) != -1) {
                    zout.write(buffer, 0, len);
                }
                zout.closeEntry();
                zin.closeEntry();
            }
        }
    }

    /**
     * @param sheet
     * @param srcCell
     * @param columnBegin
     * @param columnEnd
     * @param buocNhay
     */
    private static void copySrcCellToRange(Sheet sheet, Cell srcCell, int columnBegin, int columnEnd, int buocNhay) {
        String originalFormula = srcCell.getCellFormula();
        int srcRowIndex = srcCell.getRowIndex();
        int srcColIndex = srcCell.getColumnIndex();
        for (int i = columnBegin; i < columnEnd; i = i + buocNhay) {
            Cell sauHangSanPhamCuoiThuI = sheet.getRow(srcRowIndex).getCell(i);
            String newFormula = shiftFormulaColumns(originalFormula, i - srcColIndex);
            sauHangSanPhamCuoiThuI.setCellType(CellType.FORMULA);
            sauHangSanPhamCuoiThuI.setCellFormula(newFormula);
        }
    }

    /**
     * copy nhiều range tại các hàng trong rowsToCopy và nằm trong khảng cột startCol đến cột endCol giữ nguyên giá trị và giá trị gốc công thức từ sheet của file excel nguồn sang file excel đích
     *
     * @param srcSheet
     * @param dstSheet
     * @param rowsToCopy
     * @param startCol
     * @param endCol
     */
    public static void saoChepCacHangTrongVungCotTuFileMauVaoFileDich(Sheet srcSheet, Sheet dstSheet, int[] rowsToCopy, int startCol, int endCol, XSSFWorkbook destWorkbook) {
//        Sheet srcSheet = sourceWorkbook.getSheetAt(0);
//        Sheet dstSheet = destWorkbook.getSheetAt(0);
//        // Các hàng cần copy (zero-based): hàng 7 -> index 6, hàng 10 -> index 9
//        int[] rowsToCopy = { 6, 9 };
//        // Phạm vi cột: AN -> 39 (0-based), BA -> 52 (0-based)
//        int startCol = 39, endCol = 52;

        // Sao chép từng ô trong hai hàng
        for (int rowIndex : rowsToCopy) {
            Row srcRow = srcSheet.getRow(rowIndex);
            if (srcRow == null) continue;
            Row dstRow = dstSheet.getRow(rowIndex);
            if (dstRow == null) dstRow = dstSheet.createRow(rowIndex);

            for (int col = startCol; col <= endCol; col++) {
                Cell srcCell = srcRow.getCell(col);
                if (srcCell == null) continue;
                Cell dstCell = dstRow.createCell(col);

                // Sao chép giá trị hoặc công thức
                switch (srcCell.getCellType()) {
                    case STRING:
                        dstCell.setCellValue(srcCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(srcCell)) {
                            dstCell.setCellValue(srcCell.getDateCellValue());
                        } else {
                            dstCell.setCellValue(srcCell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
                        dstCell.setCellValue(srcCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        dstCell.setCellFormula(srcCell.getCellFormula());
                        break;
                    case BLANK:
                        dstCell.setBlank();
                        break;
                    default:
                        // Những loại khác (nếu có) có thể thêm xử lý tương tự
                        dstCell.setCellValue(srcCell.toString());
                        break;
                }

                // Sao chép định dạng ô (style)
                CellStyle newStyle = destWorkbook.createCellStyle();
                newStyle.cloneStyleFrom(srcCell.getCellStyle());
                dstCell.setCellStyle(newStyle);
            }

            // Sao chép các vùng hợp nhất (merged regions) có liên quan
            int mergedCount = srcSheet.getNumMergedRegions();
            for (int i = 0; i < mergedCount; i++) {
                CellRangeAddress region = srcSheet.getMergedRegion(i);
                int firstRow = region.getFirstRow();
                int lastRow = region.getLastRow();
                int firstCol = region.getFirstColumn();
                int lastCol = region.getLastColumn();
                // Kiểm tra xem vùng hợp nhất có nằm hoàn toàn trong AN7:BA7 hoặc AN10:BA10 không
                if ((firstRow == rowIndex && lastRow == rowIndex && firstCol >= startCol && lastCol <= endCol)) {
                    dstSheet.addMergedRegion(
                            new CellRangeAddress(firstRow, lastRow, firstCol, lastCol)
                    );
                }
            }
        }


    }

    /**
     * copy srcCell sang destCell, nếu cell là công thức sẽ update công thức
     *
     * @param srcCell      cell gốc
     * @param destCell     cell cần set giá trị copy từ cell gốc
     * @param shiftColumns
     */
    private static void copyCellWithFormulaUpdate(Cell srcCell, Cell destCell, int shiftColumns) {
        if (srcCell == null || destCell == null) {
            return;
        }
        // gán kiểu của cell gốc cho cell mới
        destCell.setCellStyle(srcCell.getCellStyle());
        switch (srcCell.getCellType()) {
            case STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case FORMULA:
                String formula = srcCell.getCellFormula();
                StringBuilder updatedFormula = new StringBuilder(updateFormula(formula, shiftColumns, srcCell.getRowIndex()));
                updatedFormula = new StringBuilder(updatedFormula.toString().replaceAll("SUN", "SUM"));

                char[] formulaArr = formula.toCharArray();
                char[] updatedFormulaArr = updatedFormula.toString().toCharArray();

                // đổi công thức cũ và mới sang dạng list rồi thêm khóa $ của công thức cũ sang các vị trí tương tự ở công thức mới
                List<String> formulaList = new ArrayList<>();
                List<String> updatedFormulaList = new ArrayList<>();

                // Duyệt qua từng ký tự trong chuỗi và thêm vào danh sách
                for (char ch : formulaArr) {
                    formulaList.add(String.valueOf(ch));
                }

                // Duyệt qua từng ký tự trong chuỗi và thêm vào danh sách
                for (char ch : updatedFormulaArr) {
                    updatedFormulaList.add(String.valueOf(ch));
                }

                for (int i = 0; i < formulaList.size(); i++) {
                    String old = formulaList.get(i);
                    if (old.equalsIgnoreCase("$")) {
                        updatedFormulaList.add(i, "$");
                    }
                }

                updatedFormula = new StringBuilder();
                for (String s : updatedFormulaList) {
                    updatedFormula.append(s);
                }

                destCell.setCellFormula(updatedFormula.toString());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            default:
                break;
        }
    }

    private static String updateFormula(String formula, int shiftColumns, int rowIndex) {
        StringBuilder updatedFormula = new StringBuilder();
        int length = formula.length();

        for (int i = 0; i < length; i++) {
            // lấy chữ cái tại vị trí i
            char c = formula.charAt(i);
            // nếu là chữ thông thương và ký tự khóa công thức $
            if (Character.isLetter(c) || c == '$') {

                StringBuilder reference = new StringBuilder();
                boolean isColumnAbsolute = false;
                boolean isRowAbsolute = false;


                if (c == '$') {
                    isColumnAbsolute = true;// khóa cột
                    reference.append(c);
                    i++;
                    c = formula.charAt(i);
                }

                // lấy các chữ cái đằng sau ký tự khóa $ và tăng i đến khi hết kí tự tức là hết tham chiếu đến 1 cell
                while (i < length && Character.isLetter(formula.charAt(i))) {
                    reference.append(formula.charAt(i));
                    i++;
                }

                // nếu tiếp tục còn khóa thì là khóa hàng
                if (i < length && formula.charAt(i) == '$') {
                    isRowAbsolute = true;
                    reference.append(formula.charAt(i));
                    i++;
                }

                while (i < length && Character.isDigit(formula.charAt(i))) {
                    reference.append(formula.charAt(i));
                    i++;
                }

                String column = reference.toString().replaceAll("[^A-Z]", "");
                String row = reference.toString().replaceAll("[^0-9]", "");

                if (!isColumnAbsolute) {
                    int columnIndex = columnToIndex(column) + shiftColumns;
                    updatedFormula.append(indexToColumn(columnIndex));
                } else {
                    updatedFormula.append(column);
                }

                if (!isRowAbsolute && !row.isEmpty()) {
                    updatedFormula.append(row);
                } else {
                    updatedFormula.append(row);
                }

                i--; // Adjust for the increment in the loop
            } else {
                updatedFormula.append(c);
            }
        }
        return updatedFormula.toString();
    }


    private static int columnToIndex(String column) {
        int index = 0;
        for (int i = 0; i < column.length(); i++) {
            index = index * 26 + (column.charAt(i) - 'A' + 1);
        }
        return index - 1;
    }

    private static String indexToColumn(int index) {
        StringBuilder column = new StringBuilder();
        while (index >= 0) {
            column.insert(0, (char) ('A' + (index % 26)));
            index = index / 26 - 1;
        }
        return column.toString();
    }

    /**
     * copy công thức từ srcCell vào destCell và thay đổi công thức theo hàng
     * của destCell cho phù hợp dựa vào khoảng cách hàng giữa 2 cell bằng tham số shiftRows
     *
     * @param srcCell   cell gốc
     * @param destCell  cell copy từ cell gốc
     * @param shiftRows khảng cách hàng giữa cell gốc và cell copy
     */
    private static void copyRowCellWithFormulaUpdate(Cell srcCell, Cell destCell, int shiftRows) {
        destCell.setCellStyle(srcCell.getCellStyle());
        switch (srcCell.getCellType()) {
            case STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case FORMULA:
                String formula = srcCell.getCellFormula();
                StringBuilder updatedFormula = new StringBuilder(updateRowFormula(formula, shiftRows));
                updatedFormula = new StringBuilder(updatedFormula.toString().replaceAll("SUN", "SUM"));


                destCell.setCellFormula(updatedFormula.toString());
                break;
            case BLANK:
                destCell.setBlank();
                break;
            default:
                break;
        }
    }


    private static String updateRowFormula(String formula, int shiftRows) {
        StringBuilder updatedFormula = new StringBuilder();
        int length = formula.length();
        boolean isAbsoluteColumn = false;
        boolean isAbsoluteRow = false;

        for (int i = 0; i < length; i++) {
            char c = formula.charAt(i);
            if (c == '$') {
                if (i + 1 < length && Character.isLetter(formula.charAt(i + 1))) {
                    isAbsoluteColumn = true;
                    updatedFormula.append(c);
                } else if (i + 1 < length && Character.isDigit(formula.charAt(i + 1))) {
                    isAbsoluteRow = true;
                    updatedFormula.append(c);
                }
            } else if (Character.isLetter(c)) {
                StringBuilder column = new StringBuilder();
                while (i < length && Character.isLetter(formula.charAt(i))) {
                    column.append(formula.charAt(i));
                    i++;
                }
                if (isAbsoluteColumn) {
                    updatedFormula.append(column.toString());
                } else {
                    updatedFormula.append(column.toString());
                }
                isAbsoluteColumn = false;
                i--; // Adjust for the increment in the loop
            } else if (Character.isDigit(c)) {
                StringBuilder row = new StringBuilder();
                while (i < length && Character.isDigit(formula.charAt(i))) {
                    row.append(formula.charAt(i));
                    i++;
                }
                int rowIndex = Integer.parseInt(row.toString());
                if (!isAbsoluteRow) {
                    rowIndex += shiftRows;
                }
                updatedFormula.append(rowIndex);
                isAbsoluteRow = false;
                i--; // Adjust for the increment in the loop
            } else {
                updatedFormula.append(c);
            }
        }
        return updatedFormula.toString();
    }

    // formula: ví dụ "SUM(Z9:AA12)"
    public static String shiftFormulaColumns(String formula, int shift) {
        // Pattern bắt được tham chiếu dạng A1, Z9, AA12, ...
        Pattern p = Pattern.compile("([A-Z]+)(\\d+)");
        Matcher m = p.matcher(formula);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String colStr = m.group(1);           // "Z" hoặc "AA"
            String rowStr = m.group(2);           // "9" hoặc "12"
            int colIndex = CellReference.convertColStringToIndex(colStr);
            int newColIndex = colIndex + shift;   // dịch số cột
            String newColStr = CellReference.convertNumToColString(newColIndex);
            m.appendReplacement(sb, newColStr + rowStr);
        }
        m.appendTail(sb);
        return sb.toString();
    }

    /**
     * Copy công thức từ srcCell ra một vùng (columns x rows).
     *
     * @param sheet       sheet đích (cùng sheet hoặc khác sheet nếu cần).
     * @param srcCell     ô nguồn chứa công thức (hoặc giá trị).
     * @param columnBegin cột bắt đầu (inclusive, zero-based).
     * @param columnEnd   cột kết thúc (exclusive, zero-based).
     * @param colStep     bước nhảy cột (ví dụ 1, 2...).
     * @param rowBegin    hàng bắt đầu (inclusive, zero-based).
     * @param rowEnd      hàng kết thúc (exclusive, zero-based).
     * @param rowStep     bước nhảy hàng.
     * @param rowMajor    nếu true -> lặp hàng ngoài, cột trong (row then col). Nếu false -> cột ngoài, hàng trong.
     */
    public static void copySrcCellToRange(Sheet sheet, Cell srcCell,
                                          int columnBegin, int columnEnd, int colStep,
                                          int rowBegin, int rowEnd, int rowStep,
                                          boolean rowMajor) {

        if (sheet == null || srcCell == null) return;

        String originalFormula = (srcCell.getCellType() == CellType.FORMULA)
                ? srcCell.getCellFormula() : null;

        int srcRowIndex = srcCell.getRowIndex();
        int srcColIndex = srcCell.getColumnIndex();

        // tìm merged region chứa cell gốc (nếu có)
        CellRangeAddress srcMerged = findMergedRegion(sheet, srcRowIndex, srcColIndex);

        if (rowMajor) {
            for (int r = rowBegin; r <= rowEnd; r += rowStep) {
                for (int c = columnBegin; c <= columnEnd; c += colStep) {
                    if (r == srcRowIndex && c == srcColIndex) continue;

                    Cell target = getOrCreateCell(sheet, r, c);

//                    if (srcCell.getCellStyle().getAlignment() != null){
//                        target.getCellStyle().setAlignment(srcCell.getCellStyle().getAlignment());
//                    }
                    // gán kiểu của cell gốc cho cell mới
                    if (srcCell.getCellStyle() != null) {
                        target.setCellStyle(srcCell.getCellStyle());
                    }

                    // copy style
//                    target.setCellStyle(srcCell.getCellStyle());

                    int colShift = c - srcColIndex;
                    int rowShift = r - srcRowIndex;

                    if (originalFormula != null) {
                        String newFormula = shiftFormulaA1(originalFormula, colShift, rowShift);
                        target.setCellFormula(newFormula);
                    } else {
                        copyCellValue(srcCell, target);
                    }

                    // nếu cell gốc thuộc vùng merged => merge tương ứng
                    if (srcMerged != null) {
                        addMergedRegionSafe(sheet, srcMerged, rowShift, colShift);
                    }
                }
            }
        } else {
            for (int c = columnBegin; c <= columnEnd; c += colStep) {
                for (int r = rowBegin; r <= rowEnd; r += rowStep) {
                    if (r == srcRowIndex && c == srcColIndex) continue;

                    Cell target = getOrCreateCell(sheet, r, c);
                    // gán kiểu của cell gốc cho cell mới
                    if (srcCell.getCellStyle() != null) {
                        target.setCellStyle(srcCell.getCellStyle());
                    }

                    int colShift = c - srcColIndex;
                    int rowShift = r - srcRowIndex;

                    if (originalFormula != null) {
                        String newFormula = shiftFormulaA1(originalFormula, colShift, rowShift);
                        target.setCellFormula(newFormula);
                    } else {
                        copyCellValue(srcCell, target);
                    }

                    if (srcMerged != null) {
                        addMergedRegionSafe(sheet, srcMerged, rowShift, colShift);
                    }
                }
            }
        }
    }

    // ===================== helper ======================

    /**
     * tìm merged region chứa cell (nếu có)
     */
    private static CellRangeAddress findMergedRegion(Sheet sheet, int row, int col) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress region : mergedRegions) {
            if (region.isInRange(row, col)) {
                return region;
            }
        }
        return null;
    }

    /**
     * thêm vùng merge tương ứng (dịch theo hàng & cột), tránh trùng
     */
    private static void addMergedRegionSafe(Sheet sheet, CellRangeAddress srcMerged, int rowShift, int colShift) {
        int firstRow = srcMerged.getFirstRow() + rowShift;
        int lastRow = srcMerged.getLastRow() + rowShift;
        int firstCol = srcMerged.getFirstColumn() + colShift;
        int lastCol = srcMerged.getLastColumn() + colShift;

        CellRangeAddress newMerged = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        // tránh thêm trùng vùng merge
        for (CellRangeAddress existing : sheet.getMergedRegions()) {
            if (existing.formatAsString().equals(newMerged.formatAsString())) {
                return;
            }
        }

        sheet.addMergedRegion(newMerged);
    }

    /**
     * tạo cell nếu chưa có
     */
    private static Cell getOrCreateCell(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        if (cell == null) cell = row.createCell(colIdx);
        return cell;
    }

    /**
     * copy giá trị cơ bản nếu không có công thức
     */
    private static void copyCellValue(Cell src, Cell target) {
        if (src == null || target == null) return;
        switch (src.getCellType()) {
            case STRING:
                target.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                target.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                target.setCellValue(src.getBooleanCellValue());
                break;
            case BLANK:
                target.setBlank();
                break;
            default:
                break;
        }
    }

    /**
     * dịch công thức A1-style theo hàng/cột, tôn trọng $
     */
    public static String shiftFormulaA1(String formula, int colShift, int rowShift) {
        if (formula == null || formula.isEmpty()) return formula;
        Pattern p = Pattern.compile("((?:'[^']+'|[A-Za-z0-9_]+)!|)(\\$?)([A-Z]+)(\\$?)(\\d+)");
        Matcher m = p.matcher(formula);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String sheetPrefix = m.group(1) == null ? "" : m.group(1);
            String colDollar = m.group(2);
            String colLetters = m.group(3);
            String rowDollar = m.group(4);
            String rowNumberStr = m.group(5);

            // cột
            String newCol = colLetters;
            if (!"$".equals(colDollar) && colShift != 0) {
                int colIndex = CellReference.convertColStringToIndex(colLetters);
                int newColIndex = Math.max(0, colIndex + colShift);
                newCol = CellReference.convertNumToColString(newColIndex);
            }

            // hàng
            String newRowStr = rowNumberStr;
            if (!"$".equals(rowDollar) && rowShift != 0) {
                int rowNum = Integer.parseInt(rowNumberStr);
                int newRow = Math.max(1, rowNum + rowShift);
                newRowStr = Integer.toString(newRow);
            }

            String replacement = sheetPrefix + colDollar + newCol + rowDollar + newRowStr;
            m.appendReplacement(sb, Matcher.quoteReplacement(replacement));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    /**
     * xóa các hợp nhất ô trong vùng chỉ định mà không giữ lại giá trị các ô bên phải và bên dưới
     *
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public static void unmergeRegionsInRange(Sheet sheet,
                                             int firstRow, int lastRow,
                                             int firstCol, int lastCol) {
        // Duyệt ngược để tránh thay đổi index khi remove
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress merged = sheet.getMergedRegion(i);

            // Kiểm tra có giao nhau với vùng chỉ định không
            boolean intersects = !(merged.getFirstRow() > lastRow
                    || merged.getLastRow() < firstRow
                    || merged.getFirstColumn() > lastCol
                    || merged.getLastColumn() < firstCol);

            if (intersects) {
                sheet.removeMergedRegion(i);
            }
        }
    }

    /**
     * xóa các hợp nhất ô trong vùng chỉ định mà giữ lại giá trị tại tất cả các ô
     *
     * @param sheet
     * @param firstRow
     * @param lastRow
     * @param firstCol
     * @param lastCol
     */
    public static void unmergeAndFillCellsInRange(Sheet sheet,
                                                  int firstRow, int lastRow,
                                                  int firstCol, int lastCol) {
        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
            CellRangeAddress merged = sheet.getMergedRegion(i);

            boolean intersects = !(merged.getFirstRow() > lastRow
                    || merged.getLastRow() < firstRow
                    || merged.getFirstColumn() > lastCol
                    || merged.getLastColumn() < firstCol);

            if (!intersects) continue;

            int r1 = merged.getFirstRow();
            int c1 = merged.getFirstColumn();

            Row row = sheet.getRow(r1);
            Cell topLeft = (row == null) ? null : row.getCell(c1);

            Object value = null;
            CellType type = null;
            CellStyle style = null;

            if (topLeft != null) {
                type = topLeft.getCellType();
                style = topLeft.getCellStyle();

                switch (type) {
                    case STRING:
                        value = topLeft.getStringCellValue();
                        break;
                    case NUMERIC:
                        value = topLeft.getNumericCellValue();
                        break;
                    case BOOLEAN:
                        value = topLeft.getBooleanCellValue();
                        break;
                    case FORMULA:
                        value = topLeft.getCellFormula();
                        break;
                    case ERROR:
                        value = topLeft.getErrorCellValue(); // byte
                        break;
                    case BLANK:
                    default:
                        value = null;
                        break;
                }
            }

            // Remove merged region trước khi điền từng ô
            sheet.removeMergedRegion(i);

            for (int r = r1; r <= merged.getLastRow(); r++) {
                Row rr = sheet.getRow(r);
                if (rr == null) rr = sheet.createRow(r);
                for (int c = c1; c <= merged.getLastColumn(); c++) {
                    Cell cell = rr.getCell(c);
                    if (cell == null) cell = rr.createCell(c);

                    if (style != null) cell.setCellStyle(style);

                    if (value != null && type != null) {
                        switch (type) {
                            case STRING:
                                cell.setCellValue((String) value);
                                break;
                            case NUMERIC:
                                cell.setCellValue(((Number) value).doubleValue());
                                break;
                            case BOOLEAN:
                                cell.setCellValue((Boolean) value);
                                break;
                            case FORMULA:
                                cell.setCellFormula((String) value);
                                break;
                            case ERROR:
                                if (value instanceof Byte) {
                                    cell.setCellErrorValue((Byte) value);
                                } else if (value instanceof Number) {
                                    cell.setCellErrorValue(((Number) value).byteValue());
                                }
                                break;
                            default:
                                cell.setBlank();
                                break;
                        }
                    } else {
                        cell.setBlank();
                    }
                }
            }
        }
    }

}