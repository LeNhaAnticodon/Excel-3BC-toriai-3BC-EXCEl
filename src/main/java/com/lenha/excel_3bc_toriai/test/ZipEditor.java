package com.lenha.excel_3bc_toriai.test;


import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class ZipEditor {
    public static final Path copyFile = Paths.get("C:\\Users\\HuanTech PC\\Desktop\\testConvertTo3BC\\NC_Excel_to_3BC.zip");
    public static final String zipFilePath = copyFile.toAbsolutePath().toString();
    public static final String fileNameToEdit = "Excel_to_3BC/Products/Product.dat";
    public static final String newText = """
            FILE_VERSION=1\r
            PRODUCT=0001\r
            CRWH-27,,H,2000,1000,55,80,80,0,19256,1,0,0.0,0.0,0,0,0,0,0,0,20.900,963.9,0,,,,1\r
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
            PRODUCT=0002\r
            CRWH-26,,H,1000,1000,60,80,80,0,19255,1,0,0.0,0.0,0,0,0,0,0,0,16.900,963.9,0,,,,1\r
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
            PRODUCT=0003\r
            CRWH-25,,H,1500,1000,32,45,80,0,24589,1,0,0.0,0.0,0,0,0,0,0,0,10.800,1230.6,0,,,,1\r
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
            PRODUCT=0004\r
            CRWH-24,,H,1000,500,50,70,80,0,19254,1,0,0.0,0.0,0,0,0,0,0,0,9.300,963.8,0,,,,1\r
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
            PRODUCT=0005\r
            CRWH-23,,H,1750,900,50,80,80,0,24589,3,0,0.0,0.0,0,0,0,0,0,0,18.000,1230.6,0,,,,1\r
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
            END
            """;
    String newFolderName = "Excel_to_3BC";  // Tên thư mục mới

    public static void main(String[] args) throws FileNotFoundException {
        // Đọc file mẫu từ resources rồi copy file ra địa chỉ của copyFile
        try (InputStream sourceFile = ZipEditor.class.getResourceAsStream("/com/lenha/excel_3bc_toriai/sampleFiles/NC_Excel_to_3BC.zip")) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }

            File copy = copyFile.toFile();
            if (copy.exists()) {
                if (copy.delete()) {
                    System.out.println("File đã được xóa thành công.");
                } else {
                    System.out.println("Xóa file thất bại.");
                }
            } else {
                System.out.println("File không tồn tại.");
            }

            Files.copy(sourceFile, copyFile);
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            // Đọc tệp ZIP
            FileInputStream fis = new FileInputStream(zipFilePath);
            ZipInputStream zis = new ZipInputStream(fis);
            ZipEntry entry;

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ZipOutputStream zos = new ZipOutputStream(baos, Charset.forName("MS932"));

            while ((entry = zis.getNextEntry()) != null) {
                // Nếu là file cần chỉnh sửa
                System.out.println(entry.getName());
                if (entry.getName().equals(fileNameToEdit)) {
                    ByteArrayOutputStream tempBaos = new ByteArrayOutputStream();
                    byte[] buffer = new byte[1024];
                    int len;
                    // đọc các byte data của file cần chỉnh này lưu vào buffer
                    // zis.read(buffer) trả về độ dài của data đã đọc
                    // các đoạn buffer được tempBaos ghi lại để tổng hợp thành nội dung hoàn chỉnh
                    while ((len = zis.read(buffer)) > 0) {
                        tempBaos.write(buffer, 0, len);
                    }

                    // lấy ra dữ liệu cũ của file đã đọc trong tempBaos
                    String oldContent = tempBaos.toString(StandardCharsets.UTF_8);
                    System.out.println(oldContent);
                    // tạo đoạn dữ liệu mới muốn ghi đè vào file đang lặp
                    String newContent = newText;

                    // Viết file đã chỉnh sửa vào tệp ZIP mới
                    // gán entry của file đang lặp cho trình ghi zip
                    zos.putNextEntry(new ZipEntry(entry.getName()));
                    // ghi đè dữ liệu của file đang lặp bằng newContent
                    zos.write(newContent.getBytes());

                    zos.closeEntry();
                } else {
                    // Copy các file không chỉnh sửa vào tệp ZIP mới
                    // gán entry của file đang lặp cho trình ghi zip
                    zos.putNextEntry(new ZipEntry(entry.getName()));
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

            // Ghi tệp ZIP mới vào đĩa
            FileOutputStream fos = new FileOutputStream(zipFilePath);
            baos.writeTo(fos);
            fos.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

