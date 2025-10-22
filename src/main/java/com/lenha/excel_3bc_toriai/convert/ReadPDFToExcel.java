package com.lenha.excel_3bc_toriai.convert;

import com.lenha.excel_3bc_toriai.model.ExcelFile;
import com.opencsv.CSVWriter;
import javafx.collections.ObservableList;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.charset.Charset;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.TimeoutException;

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

    // tên file chl sẽ tạo được ghi trong phần 工事名, chưa bao gồm loại vật liệu
    private static String fileChlName = "";

    // link của file pdf
    private static String pdfPath = "";

    // link thư mục của file excel xlsx sẽ tạo
    private static String xlsxExcelPath = "";
    // link thư mục của file excel csv sẽ tạo
    private static String csvExcelDirPath = "";
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
    private static  double kouzaiChouGoukei = 0;
    private static double seiHinChouGoukei = 0;

    /**
     * chuyển đổi pdf tính vật liệu thành các file chl theo từng vật liệu khác nhau
     * @param filePDFPath link file pdf
     * @param fileChlDirPath link thư mục chứa file chl sẽ tạo
     * @param excelFileNames list chứa danh sách các file chl đã tạo
     */
    public static void convertPDFToExcel(String filePDFPath, String fileChlDirPath, ObservableList<ExcelFile> excelFileNames) throws FileNotFoundException, TimeoutException, IOException {
/*        excelFileNames.add(new ExcelFile("test.", "", 0, 0));
        fileName = "test.sysc2";
        throw new TimeoutException();*/

        // xóa danh sách cũ trước khi thực hiện, tránh bị ghi chồng lên nhau
        excelFileNames.clear();

        // lấy địa chỉ file pdf
        pdfPath = filePDFPath;
        // lấy đi chỉ thư mục chứa file excel
//        csvExcelDirPath = fileCSVDirPath;
        // lấy đi chỉ thư mục chứa file excel csv
        csvExcelDirPath = fileChlDirPath;
        // lấy đi chỉ thư mục chứa chl
        chlDirPath = fileChlDirPath;

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

            if (i > 1) {
                String KouSyuNameBefore = extractValue(kakuKouSyuList.get(i - 1), "法:", "梱包");

                if (KouSyuName.equals(KouSyuNameBefore)) {
                    kakuKouSyuList.set(i - 1, kakuKouSyuList.get(i - 1).concat(kakuKouSyuList.get(i)));
                    kakuKouSyuList.remove(i);
                    i--;
                    kakuKouSyuListSize--;
                }
            }
        }

        // tạo số thứ tự khi ghi tên là thời gian ở ô tên trong file chl để tránh trùng thời gian
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
        }

    }

    /**
     * lấy toàn bộ text của file pdf
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
     * lấy các thông tin của đơn và ghi vào các biến nhớ toàn cục
     * các thông tin nằm trong vùng xác định, dùng hàm extractValue để lấy
     * @param header text chứa thông tin
     */
    private static void getHeaderData(String header) {
        String nouKi = extractValue(header, "期[", "]");
        String[] nouKiArr = nouKi.split("/");
        shortNouKi = nouKiArr[1] + "/" + nouKiArr[2];

        kouJiMe = extractValue(header, "考[", "]");
        kyakuSakiMei = extractValue(header, "客先名[", "]");
        fileChlName = extractValue(header, "工事名[", "]");

        System.out.println(shortNouKi + " : " + kouJiMe + " : " + kyakuSakiMei);
    }

    /**
     * lấy thông số đầy đủ của vật liệu, tên vật liệu, mã vật liệu, 3 size của vật liệu và ghi vào biến toàn cục
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
     * @return list các map các đoạn tính vật liệu chứa key cũng là map chỉ có 1 cặp có key là chiều dài bozai, value là số lượng bozai,
     * mỗi phần tử của list sẽ tạo 1 file
     * còn value của kaKouPairs cũng là map chứa các cặp key là mảng 2 phần tử gồm tên và chiều dài sản phẩm, value là số lượng sản phẩm
     */
    private static List<Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>>> getToriaiData(String[] kakuKakou) throws TimeoutException {
        // reset lại danh sách list file
        fileList.clear();

        // tạo map
        Map<Map<StringBuilder, Integer>, Map<StringBuilder[], Integer>> kaKouPairs = new LinkedHashMap<>();

        // nếu không có thông tin thì thoát
        if (kakuKakou == null) {
            return fileList;
        }

        // lặp qua các đoạn bozai và thêm chúng vào map chứa toàn bộ thông tin vật liệu
        for (int i = 1; i < kakuKakou.length; i++) {
            // lấy kirirosu tại lần 1
            if (i == 1) {
                kirirosu = extractValue(kakuKakou[i], "切りﾛｽ設定:", "mm");
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
                if (line.contains("鋼材長:") && line.contains("本数:")) {
                    String kouZaiChou = extractValue(line, "鋼材長:", "mm");
                    String honSuu = extractValue(line, "本数:", " ").split(" ")[0];
                    // mẫu định dạng "#.##". Mẫu này chỉ hiển thị phần thập phân nếu có, và tối đa là 2 chữ số thập phân.
                    DecimalFormat df = new DecimalFormat("#.##");
                    kouZaiChouPairs.put(new StringBuilder().append(df.format(Double.parseDouble(kouZaiChou.trim()))), convertStringToIntAndMul(honSuu.trim(), 1));
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
                    String length = meiSyouLengths[meiSyouLengths.length - 1];

                    // thêm tên và chiều dài vào mảng với chiều dài x 100
                    StringBuilder[] nameAndLength = {new StringBuilder().append(name), new StringBuilder().append(convertStringToIntAndMul(length.trim(), 100))};

                    // lấy số lượng sản phẩm
                    String meiSyouHonSuu = extractValue(line, "mm x", "本").trim();
                    // thêm cặp tên + chiều dài và số lượng vào map
                    meiSyouPairs.put(nameAndLength, convertStringToIntAndMul(meiSyouHonSuu, 1));
                }
            }

            // thêm 2 map chứa thông tin vật liệu vào map gốc
            kaKouPairs.put(kouZaiChouPairs, meiSyouPairs);
        }

/*        // in thông tin vật liệu
        kaKouPairs.forEach((kouZaiChouPairs, meiSyouPairs) -> {
            kouZaiChouPairs.forEach((key, value) -> System.out.println("\n" + key.toString() + " : " + value));
            meiSyouPairs.forEach((key, value) -> System.out.println(key[0].toString() + " " + key[1].toString() + " : " + value));
        });*/


//        if (checkRowNum(kaKouPairs) > 99) {
            divFile(kaKouPairs);
//        }

        System.out.println(rowToriAiNum);
        System.out.println("\n" + kirirosu);

        // trả về list các map kết quả để ghi vào file chl sysc2
        return fileList;
    }

    /**
     * chia file theo số dòng sản phẩm
     * nếu vượt quá 99 dòng sẽ chia thành 2 file, file 1 là dưới 99 dòng rồi thêm nó vào list file, file 2 là phần còn lại
     * tiếp tục gọi lại hàm(đệ quy) và truyền file 2 vào và thực hiện tương tự đến khi không cần chia thành file 2 nữa tức
     * file 2 có size = 0 thì dừng
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
        String excelPath = csvExcelDirPath + "\\" + fileName;

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
        String csvPath = csvExcelDirPath + "\\" + fileName;
        System.out.println("dir path: " + csvExcelDirPath);
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
     * @param kaKouPairs   map chứa tính vật liệu
     * @param timePlus     thời gian hoặc chỉ số cộng thêm vào ô time để tránh bị trùng tên  time giữa các file
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
        System.out.println("dir path: " + csvExcelDirPath);
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
     * @param text đoạn văn bản chứa thông tin tìm kiếm
     * @param startDelimiter đoạn text phía trước vùng cần tìm
     * @param endDelimiter đoạn text phía sau vùng cần tìm
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

}
