package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import com.lenha.excel_3bc_toriai.convert.ReadPDFToExcel;
import com.lenha.excel_3bc_toriai.test.ZipEditor;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class Write3BC {
    private static String _3bcDirPath;
    private static String doanBatDauFileSanPham = """
            FILE_VERSION=1\r
            """;
    private static String tieuDeSttSanPham = """
            PRODUCT=
            """.stripIndent();

  /*  private static String codeChungCuaCacSanPham = "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r" +
            "    MID=0,  \r";*/

    public static final String codeChungCuaCacSanPham = """
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
                MID=0,  \r
            """;
    private static String ketThucFileSanPham = """
            END\r
            """;


    public static Path copyFile = null;
    private static String oldFolderName = "Excel_to_3BC";

    public static void write3BCFile(String file3bcDirPath, Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets) throws IOException {
        String[] thongTinDonHang = null;

        // lấy ra thông tin đơn hàng bằng cách lấy thông tin ở sheet đầu tiên
        for (Map.Entry<String[], List<Map.Entry<Double, Integer>>> entry : toriaiSheets.entrySet()) {
            thongTinDonHang = entry.getKey();
            break;
        }
        if (thongTinDonHang == null) {
            throw new IOException("Không có thông tin đơn hàng");
        }

        // lấy mã đơn
        String maDon = thongTinDonHang[6].trim();
        // nếu mã đơn hàng dài hơn 30 kí tự thì cắt cho còn 30, vì khi nhập tên trên ô nhập của máy 3bc chỉ cho tối đa 30 kí tự
        if (maDon.length() > 30) {
            maDon = maDon.substring(0, 30);
        }

        // lấy đi chỉ thư mục chứa 3bc
        _3bcDirPath = file3bcDirPath;
        // Đọc file mẫu từ resources rồi copy file ra địa chỉ của copyFile
        try (InputStream sourceFile = ZipEditor.class.getResourceAsStream("/com/lenha/excel_3bc_toriai/sampleFiles/NC_Excel_to_3BC.zip")) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }


            // tạo tên file 3bc theo đúng cấu trúc
            String copyFileName = "\\NC_" + maDon + ".zip";

            copyFile = Paths.get(_3bcDirPath + copyFileName);

            // lấy địa chỉ file sẽ được copy
            File copy = copyFile.toFile();

            // nếu file đã tồn tại thì xóa nó
            if (copy.exists()) {
                if (copy.delete()) {
                    System.out.println("File đã được xóa thành công.");
                } else {
                    System.out.println("Xóa file thất bại.");
                }
            } else {
                System.out.println("File không tồn tại.");
            }

            // tạo file copy từ file gốc
            Files.copy(sourceFile, copyFile);
        }

        String zipFilePath = copyFile.toAbsolutePath().toString();

        // Đọc tệp ZIP của file copy
        FileInputStream fis = new FileInputStream(zipFilePath);
        ZipInputStream zis = new ZipInputStream(fis);
        ZipEntry entry;

        // ghi tệp ZIP của file copy
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ZipOutputStream zos = new ZipOutputStream(baos, Charset.forName("MS932"));
//        ZipOutputStream zos = new ZipOutputStream(baos);

        // duyệt qua các file hoặc thư mục cùng cấp trong file nén đang edit
        while ((entry = zis.getNextEntry()) != null) {
            System.out.println(entry.getName());

            String fileEditName = entry.getName();

            // Kiểm tra nếu entry thuộc thư mục hoặc file cần đổi tên, do đây là file mẫu nên cần đổi tên thư mục gốc thành tên mã đơn để đồng bộ, không làm file
            // bị lỗi khi nhập vào 3bc
            // trong trường hợp này là thư mục đầu tiên trong file nén, nó trùng với tên của file nén luôn, cũng chính là mã đơn như đã quy ước
            if (entry.getName().startsWith(oldFolderName + "/")) {
                // Đổi tên thư mục hoặc file để khi ghi lại tai đoạn zos.putNextEntry(name) sẽ ghi thành file mới với tên này
                fileEditName = maDon + entry.getName().substring(oldFolderName.length());
            }

            // tạo link file chuẩn của file sản phẩm, chỉ khi nào thư mục gốc của file nén đã được đổi tên đúng thì các file con trong thư mục đó mới có thể chỉnh sửa
            // vì các đoạn code dưới tuân theo tên này
            String fileSanPham = maDon + "/Products/Product.dat";
            // tạo link của file chứa các thông tin đơn hàng
            String fileThongTinDon = maDon + "/Koji.dat";

            // entry là file và tệp đang lặp trong file nén đang edit
            // Nếu là file cần chỉnh sửa, tức là tên entry trùng với tên file cần chỉnh sửa
            // trường hợp này là file chứa các sản phẩm
            if (fileEditName.equals(fileSanPham)) {

                // đoạn mã này chỉ có chức năng đọc lại dữ liệu của file cũ, không có tác dụng mấy nếu không cần dữ liệu cũ
                /*ByteArrayOutputStream tempBaos = new ByteArrayOutputStream();
                byte[] buffer = new byte[1024];
                int len;
                // đọc các byte data của file cần chỉnh này lưu vào buffer
                // hàm sẽ đọc data và ghi vào buffer
                // zis.read(buffer) trả về độ dài của data đã đọc
                // các đoạn buffer được tempBaos ghi lại để tổng hợp thành nội dung hoàn chỉnh
                while ((len = zis.read(buffer)) > 0) {
                    tempBaos.write(buffer, 0, len);
                }

                // lấy ra dữ liệu cũ của file đã đọc trong tempBaos
                String oldContent = tempBaos.toString(StandardCharsets.UTF_8);
                System.out.println(oldContent);*/


                // tạo dữ liệu chứa các sản phẩm của đơn
                String textSanPham = "";
                // thêm đoạn đầu file
                textSanPham = textSanPham.concat(doanBatDauFileSanPham);

                int stt = 1;
                // lấy ra thông tin các sản phẩm
                for (Map.Entry<String[], List<Map.Entry<Double, Integer>>> entry1 : toriaiSheets.entrySet()) {
                    // lấy các thông tin của vật liệu đang duyệt
                    String[] thongTin = entry1.getKey();
                    // lấy ra list các sản phẩm
                    List<Map.Entry<Double, Integer>> cacSanPham = entry1.getValue();
                    // đặt giới hạn số lượng sản phẩm
                    if (cacSanPham.size() > 9999) {
                        throw new IOException("số lượng sản phẩm vượt 999");
                    }

                    // lấy ra các size của vật liệu x 10
                    int size1 = ReadPDFToExcel.convertStringToIntAndMul(thongTin[2], 10);
                    int size2 = ReadPDFToExcel.convertStringToIntAndMul(thongTin[3], 10);
                    int size3 = ReadPDFToExcel.convertStringToIntAndMul(thongTin[4], 10);
                    int size4 = ReadPDFToExcel.convertStringToIntAndMul(thongTin[5], 10);

                    // đặt giới hạn cho các size
                    if (size1 > 99999 || size2 > 99999 || size3 > 99999 || size4 > 99999) {
                        throw new IOException("kích thước vật liệu vượt 99999");
                    }

                    // lấy ra khối lượng riêng và fomat cho hiển thị 3 chữ số phần thập phân
                    // Dùng Double.toString để tránh lỗi binary-floating representation
                    BigDecimal bd = new BigDecimal(thongTin[0]);
                    bd = bd.setScale(3, RoundingMode.DOWN); // làm tròn tới 3 chữ số thập phân
                    String khoiLuongRieng = bd.toPlainString();

                    // lặp qua các cặp chiều dài, số lượng và thêm nó vào text sản phẩm
                    for (Map.Entry<Double, Integer> sanPham : cacSanPham) {
                        // fomat số thứ tự sản phẩm về dạng số tự nhiên hiển thị  tối đa 4 chữ
                        String sttSanPham = String.format("%04d", stt);

                        // lấy chiều dài sản phẩm đã nhân với 10
                        int chieuDai = ReadPDFToExcel.convertStringToIntAndMul(sanPham.getKey().toString(), 10);
                        //đặt giới hạn chiều dài và số lượng
                        if (chieuDai > 125000) {
                            throw new IOException("chiều dài sản phẩm vượt 125000");
                        }
                        int soLuong = sanPham.getValue();
                        if (soLuong > 125000) {
                            throw new IOException("số lượng sản phẩm vượt 125000");
                        }

                        // lấy số phụ thuộc vào chiều dài, nó bằng chiều dài gốc(chưa x 10) / 2 + 1.15, làm tròn đến 1 chữ số phần thập phân
                        // phải dùng BigDecimal vì double chia bình thường sẽ có sai số
                        // Dùng Double.toString để tránh lỗi binary-floating representation
                        BigDecimal bdChieuDai = new BigDecimal(chieuDai);
                        bdChieuDai = bdChieuDai.divide(new BigDecimal(20)).add(new BigDecimal("1.15"));
                        bdChieuDai = bdChieuDai.setScale(1, RoundingMode.DOWN); // cắt (không làm tròn)
                        String soPhuThuocChieuDai = String.format(Locale.US, "%.1f", bdChieuDai.doubleValue());

                        // thêm đoạn stt sản phẩm
                        textSanPham = textSanPham.concat(tieuDeSttSanPham) + sttSanPham + "\r";

                        // thêm các thông tin sản phẩm theo cấu trúc tên,,mã vật liệu,size1,size2,size3,size4,bo góc 1, bo góc 2
                        // ,chiều dài, số lượng,0,0.0,0.0,0,0,0,0,0,0,khối lượng riêng, số phụ thuộc chiều dài,0,,,,1 + đoạn code chung cho các sản phẩm
                        // đoạn này chỉ áp dụng chi sản phẩm không tạo lỗ
                        textSanPham = textSanPham.concat(",,") + thongTin[1] + "," + size1 + "," + size2 + ","
                                + size3 + "," + size4 + ",0,0," + chieuDai + "," + soLuong
                                + ",0,0.0,0.0,0,0,0,0,0,0," + khoiLuongRieng + "," + soPhuThuocChieuDai + ",0,,,,1\r"
                                + codeChungCuaCacSanPham;

                        // tăng số thứ tự sản phẩm lên 1
                        stt++;
                    }
                }

                textSanPham = textSanPham.concat(ketThucFileSanPham);

                System.out.println("danh sách các sản phẩm: \n" + textSanPham);


                // Viết file đã chỉnh sửa vào tệp ZIP mới
                // gán entry của file đang lặp cho trình ghi zip
                zos.putNextEntry(new ZipEntry(fileEditName));
                // ghi đè dữ liệu của file đang lặp bằng textSanPham
                zos.write(textSanPham.getBytes());

                zos.closeEntry();
            }
            else if (fileEditName.equals(fileThongTinDon)) {
                // tạo dữ liệu chứa thông tin đơn hàng
                String textThongTinDon = "";

//                ZoneId zone = ZoneId.of("Asia/Bangkok"); // hoặc ZoneId.systemDefault()
                ZoneId zone = ZoneId.systemDefault();
                LocalDateTime now = LocalDateTime.now(zone);

                DateTimeFormatter dateFmt = DateTimeFormatter.ofPattern("yyyyMMdd");
                DateTimeFormatter timeFmt = DateTimeFormatter.ofPattern("HHmmss"); // HH = 24-hour (00-23), hh là định dạng 12 giờ

                String date = now.format(dateFmt); // vd: "20251022"
                String time = now.format(timeFmt); // vd: "230507" cho 23:05:07, "000507" cho 00:05:07

                // tạo thông tin đơn hàng theo định dạng
                /* mẫu là
                3 8288488                               20251006                                皆越鉄工12                              20251001  124850    沢野商会12
                1.	Vàng là tên đơn
                2.	Xanh lá cây là ngày của đơn
                3.	Xám là nơi giao hàng
                4.	Xanh da trời là ngày tạo đơn
                5.	Tím là giờ phút tạo đơn
                6.	Xanh đậm là tên khách hàng

                1, 2, 3, 6 đều có tối đa 40 kí tự, nếu chữ nhật thì cứ 1 chữ tính 2 kí tự, nghĩa là nếu toàn chữ nhật thì chỉ được tối đa 20 kí tự

                4 và 5 thì được tính 10 kí tự */
                textThongTinDon = textThongTinDon.concat(thongTinDonHang[6] + thongTinDonHang[9] +
                        thongTinDonHang[8] + date + "  " + time + "  " + "  " + thongTinDonHang[7]);

                // Viết file đã chỉnh sửa vào tệp ZIP mới
                // gán entry của file đang lặp cho trình ghi zip
                zos.putNextEntry(new ZipEntry(fileEditName));
                // ghi đè dữ liệu của file đang lặp bằng textThongTinDon
                zos.write(textThongTinDon.getBytes());

                zos.closeEntry();
            } else {
                // Copy các file không chỉnh sửa vào tệp ZIP mới
                // gán tên entry của file đang lặp cho trình ghi zip
                zos.putNextEntry(new ZipEntry(fileEditName));
                byte[] buffer = new byte[1024];
                int len;
                // lấy mảng byte đọc được từ file cũ ghi lại vào file mới
                while ((len = zis.read(buffer)) > 0) {
                    zos.write(buffer, 0, len);
                }
                zos.closeEntry();
            }
        }

        // Đóng các stream
        zis.close();
        zos.close();

        // Ghi tệp ZIP mới vào đĩa là file copy
        // baos bao gồm zos đã lấy được thông tin ở bên trên
        // sau đó ghi vào fos với đường dẫn là file copy
        FileOutputStream fos = new FileOutputStream(zipFilePath);
        baos.writeTo(fos);
        fos.close();

    }
}
