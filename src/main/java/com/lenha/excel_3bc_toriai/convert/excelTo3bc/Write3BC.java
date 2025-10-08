package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import com.lenha.excel_3bc_toriai.test.ZipEditor;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

public class Write3BC {
    private static String _3bcDirPath;
    public static Path copyFile = null;
    public static void write3BCFile(String file3bcDirPath, Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets) throws IOException {

        // lấy đi chỉ thư mục chứa 3bc
        _3bcDirPath = file3bcDirPath;
        // Đọc file mẫu từ resources rồi copy file ra địa chỉ của copyFile
        try (InputStream sourceFile = ZipEditor.class.getResourceAsStream("/com/lenha/excel_3bc_toriai/sampleFiles/NC_Excel_to_3BC.zip")) {
            if (sourceFile == null) {
                throw new IOException("File mẫu không tồn tại trong JAR ứng dụng");
            }
            copyFile = Paths.get(_3bcDirPath);

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

        try {
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

                // entry là file và tệp đang lặp trong file nén đang edit
                // Nếu là file cần chỉnh sửa, tức là tên entry trùng với tên file cần chỉnh sửa
                if (entry.getName().equals(fileNameToEdit)) {
                    ByteArrayOutputStream tempBaos = new ByteArrayOutputStream();
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

            // Ghi tệp ZIP mới vào đĩa là file copy
            // baos bao gồm zos đã lấy được thông tin ở bên trên
            // sau đó ghi vào fos với đường dẫn là file copy
            FileOutputStream fos = new FileOutputStream(zipFilePath);
            baos.writeTo(fos);
            fos.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
