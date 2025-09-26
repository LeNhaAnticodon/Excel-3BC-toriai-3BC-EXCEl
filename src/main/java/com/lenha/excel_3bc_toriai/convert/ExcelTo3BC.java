package com.lenha.excel_3bc_toriai.convert;

import com.lenha.excel_3bc_toriai.convert.excelTo3bc.ReadExcel;
import com.lenha.excel_3bc_toriai.convert.excelTo3bc.Write3BC;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelTo3BC {
    private static final List<List<Double>> LIST_CHIEU_DAI_CAC_VAT_LIEU = new ArrayList<>();
    // map chứa các tính vật liệu ở trên các sheet
    private static final ObservableList<Map<String[], Map<Double, Integer>>> toriaiSheets = FXCollections.observableArrayList();

    public static void convertExcelTo3bc(String fileExcelPath, String file3bcDirPath) {

        try {
            // đọc file excel
            ReadExcel.readExcel(fileExcelPath, toriaiSheets);

            // ghi file 3bc
            Write3BC.write3BC(file3bcDirPath);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
