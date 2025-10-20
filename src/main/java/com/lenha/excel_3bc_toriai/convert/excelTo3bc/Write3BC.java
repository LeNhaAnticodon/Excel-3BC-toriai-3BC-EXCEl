package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import com.lenha.excel_3bc_toriai.test.ZipEditor;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class Write3BC {
    private static String _3bcDirPath;
    private static String doanBatDauFileSanPham = "FILE_VERSION=1\n";
    private static String tieuDeSttSanPham = "PRODUCT=000";
    private static String codeChungCuaCacSanPham = "    0, 0, ,0,       0.0,  0,  0,    0.0,    0.0,    0.0,    0.0, 0,    0.0\r" +
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
            "    MID=0,  \r";
    private static String ketThucFileSanPham = "END\n" +
            "\"\"\"";


    public static Path copyFile = null;
    private static String oldFolderName = "Excel_to_3BC";

    public static void write3BCFile(String file3bcDirPath, Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets) throws IOException {
        String[] thongTinDonHang = null;

        // lấy ra thông tin đơn hàng bằng cách lấy thông tin ở sheet đầu tiên
        for (Map.Entry<String[], List<Map.Entry<Double, Integer>>> entry : toriaiSheets.entrySet()) {
            thongTinDonHang = entry.getKey();
            break;
        }
        if (thongTinDonHang == null){
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
                textSanPham = textSanPham.concat(doanBatDauFileSanPham);

                // lấy ra thông tin các sản phẩm
                for (Map.Entry<String[], List<Map.Entry<Double, Integer>>> entry1 : toriaiSheets.entrySet()) {
                    String[] thongTin = entry1.getKey();
                    List<Map.Entry<Double, Integer>> cacSanPham = entry1.getValue();

                }


                // Viết file đã chỉnh sửa vào tệp ZIP mới
                // gán entry của file đang lặp cho trình ghi zip
                zos.putNextEntry(new ZipEntry(fileEditName));
                // ghi đè dữ liệu của file đang lặp bằng textSanPham
                zos.write(textSanPham.getBytes());

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
