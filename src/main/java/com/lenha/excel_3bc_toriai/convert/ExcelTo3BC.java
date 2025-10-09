package com.lenha.excel_3bc_toriai.convert;

import com.lenha.excel_3bc_toriai.convert.excelTo3bc.ReadExcel;
import com.lenha.excel_3bc_toriai.convert.excelTo3bc.Write3BC;

import java.io.IOException;
import java.util.*;

public class ExcelTo3BC {
    private static final List<List<Double>> LIST_CHIEU_DAI_CAC_VAT_LIEU = new ArrayList<>();
    // map chứa thông tin các tính vật liệu của các sheet trong excel, list này sẽ dùng để tạo file 3bc
    // gồm key là mảng chứa các thông tin về vật liệu của sản phẩm, value là danh sách các sản phẩm của vật liệu
    // value phải dùng List<Map.Entry<Double, Integer>> thay vì Map<Double, Integer> bởi vì nếu key có giá trị giống nhau thì map sẽ ghi đè key cũ làm mất sản phẩm cũ
    // List<Map.Entry<Double, Integer>> thì không như vậy
    private static final Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets = new LinkedHashMap<>();

    public static boolean convertExcelTo3bc(String fileExcelPath, String file3bcDirPath) throws IOException {

        // biến kiểm tra trong danh sách các sheet tính vật liệu có sheet nào đó có vật liệu không giống với các vật liệu đã cài đặt sẵn trong chương trình không,
        // nếu có thì sẽ thay vật liệu của toàn bộ các sheet bằng bộ vật liệu tự cho trong danh sách dự phòng đã tạo khi khởi tạo chương trình
        // bộ vật liệu dự phòng lấy từ file excel VAT_LIEU_DU_PHONG.xlsx
        boolean co1VatLieuKhongTonTai = false;
        // đọc file excel
        co1VatLieuKhongTonTai = ReadExcel.readExcelFile(fileExcelPath, toriaiSheets);

        toriaiSheets.forEach((vatLieu, sanPhams) -> {
            System.out.println(Arrays.toString(vatLieu));

            for (Map.Entry<Double, Integer> sanPham : sanPhams) {
                System.out.println(sanPham.getKey() + ": " + sanPham.getValue());
            }
//                    sanPham.forEach((chieuDai, soLuong) ->{
//                        System.out.println(chieuDai + ": " + soLuong);
//                    });
        });


        // reset lại map các tính vật liệu mỗi khi chạy xong, tránh tình trạng thực hiện hàm này lần 2 thì giá trị cũ chưa bị xóa làm tăng gấp đôi giá trị
        toriaiSheets.clear();

        // ghi file 3bc
//        Write3BC.write3BCFile(file3bcDirPath, toriaiSheets);


        return co1VatLieuKhongTonTai;
    }
}
