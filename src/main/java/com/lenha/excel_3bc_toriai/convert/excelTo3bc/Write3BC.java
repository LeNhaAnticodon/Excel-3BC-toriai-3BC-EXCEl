package com.lenha.excel_3bc_toriai.convert.excelTo3bc;

import java.util.List;
import java.util.Map;

public class Write3BC {
    private static String _3bcDirPath;
    public static void write3BCFile(String file3bcDirPath, Map<String[], List<Map.Entry<Double, Integer>>> toriaiSheets) {

        // lấy đi chỉ thư mục chứa 3bc
        _3bcDirPath = file3bcDirPath;
    }
}
