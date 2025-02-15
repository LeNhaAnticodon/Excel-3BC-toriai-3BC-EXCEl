package com.lenha.excel_3bc_toriai.convert;

import java.util.ArrayList;
import java.util.List;

public class ExcelTo3BC {
    private static final List<List<Double>> LIST_CHIEU_DAI_CAC_VAT_LIEU = new ArrayList<>();

    // link của file excel
    private static String excelPath = "";
    private static String _3bcDirPath;

    public static void convertExcelTo3bc(String fileExcelPath, String file3bcDirPath) {

        // lấy địa chỉ file excel
        excelPath = fileExcelPath;

        // lấy đi chỉ thư mục chứa 3bc
        _3bcDirPath = file3bcDirPath;
    }
}
