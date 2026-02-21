package com.lenha.excel_3bc_toriai.convert.copyRowSheetToSheet;
// Dependencies: org.apache.poi:poi-ooxml
import java.io.*;
import java.util.*;
import java.util.regex.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class RowRangeCopier {

    /**
     * Copy rows [srcStartRow..srcEndRow] (1-based inclusive) from sheet srcSheetIndex
     * into sheet destSheetIndex starting at destStartRow (1-based).
     *
     * - workbook: XSSFWorkbook already opened (will be modified).
     * - srcSheetIndex, destSheetIndex: 0-based sheet indices in the workbook.
     * - srcStartRow, srcEndRow, destStartRow: 1-based row numbers (Excel style).
     *
     * Behavior:
     * - copy cell values & formulas; formulas adjusted so that relative row references
     *   (no $ before row number) are shifted by delta = destStart0 - srcStart0.
     * - copy styles (clone per style hash), comments (shallow), column widths and hidden flags
     * - copy merged regions that intersect the source block (shifted by delta)
     *
     * Limitations:
     * - does not handle external-workbook refs, named ranges, INDIRECT dynamics, or tables/pivots fully.
     */
    public static void copyRowRange(XSSFWorkbook workbook,
                                    int srcSheetIndex,
                                    int srcStartRow, // 1-based
                                    int srcEndRow,   // 1-based
                                    int destSheetIndex,
                                    int destStartRow) {

        XSSFSheet srcSheet = workbook.getSheetAt(srcSheetIndex);
        XSSFSheet destSheet = workbook.getSheetAt(destSheetIndex);

        // convert to 0-based
        int srcStart0 = Math.max(0, srcStartRow - 1);
        int srcEnd0 = Math.max(0, srcEndRow - 1);
        int destStart0 = Math.max(0, destStartRow - 1);

        int delta = destStart0 - srcStart0;
        String srcSheetName = srcSheet.getSheetName();

        // 1) copy merged regions that intersect the source block
        // We'll shift any merged region that intersects [srcStart0..srcEnd0]
        for (int i = 0; i < srcSheet.getNumMergedRegions(); i++) {
            CellRangeAddress mr = srcSheet.getMergedRegion(i);
            if (mr.getLastRow() < srcStart0 || mr.getFirstRow() > srcEnd0) {
                continue; // no intersection
            }
            // shift entire merged region by delta
            CellRangeAddress newMr = new CellRangeAddress(
                    mr.getFirstRow() + delta,
                    mr.getLastRow() + delta,
                    mr.getFirstColumn(),
                    mr.getLastColumn()
            );
            destSheet.addMergedRegion(newMr);
        }

        // 2) copy column widths and hidden state for columns used in source block
        int maxCol = 0;
        for (int r = srcStart0; r <= srcEnd0; r++) {
            Row row = srcSheet.getRow(r);
            if (row != null) maxCol = Math.max(maxCol, row.getLastCellNum());
        }
        for (int c = 0; c <= maxCol; c++) {
            try {
                destSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
                destSheet.setColumnHidden(c, srcSheet.isColumnHidden(c));
            } catch (Exception e) { /* ignore */ }
        }

        // 3) prepare style cache
        Map<Integer, XSSFCellStyle> styleMap = new HashMap<>();

        // 4) copy rows and cells
        for (int r = srcStart0; r <= srcEnd0; r++) {
            XSSFRow srcRow = srcSheet.getRow(r);
            int destRowNum = r + delta;
            // remove existing row at destRowNum if present to avoid leftover
            XSSFRow existing = destSheet.getRow(destRowNum);
            if (existing != null) destSheet.removeRow(existing);
            XSSFRow destRow = destSheet.createRow(destRowNum);

            if (srcRow == null) continue;
            destRow.setHeight(srcRow.getHeight());

            short firstCell = srcRow.getFirstCellNum();
            short lastCell = srcRow.getLastCellNum();
            if (firstCell < 0 || lastCell < 0) continue;

            for (int c = firstCell; c < lastCell; c++) {
                XSSFCell srcCell = srcRow.getCell(c);
                if (srcCell == null) continue;
                XSSFCell destCell = destRow.createCell(c);

                // copy style
                XSSFCellStyle srcStyle = srcCell.getCellStyle();
                if (srcStyle != null) {
                    int key = srcStyle.hashCode();
                    XSSFCellStyle newStyle = styleMap.get(key);
                    if (newStyle == null) {
                        newStyle = workbook.createCellStyle();
                        try { newStyle.cloneStyleFrom(srcStyle); } catch (Exception ex) {}
                        styleMap.put(key, newStyle);
                    }
                    destCell.setCellStyle(newStyle);
                }

                // copy comment shallow
                if (srcCell.getCellComment() != null) {
                    destCell.setCellComment(srcCell.getCellComment());
                }

                // copy value or adjusted formula
                if (srcCell.getCellType() == CellType.FORMULA) {
                    String formula = srcCell.getCellFormula();
                    String adjusted = adjustFormulaRowReferences(formula, srcSheetName, delta, srcStart0, srcEnd0);
                    try {
                        destCell.setCellFormula(adjusted);
                    } catch (Exception ex) {
                        // fallback: copy cached result
                        copyCachedValue(srcCell, destCell);
                    }
                } else {
                    copyCellValue(srcCell, destCell);
                }
            }
        }

        // 5) evaluate formulas to update cached values (best-effort)
        try {
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            evaluator.evaluateAll();
        } catch (Exception e) {
            // ignore evaluation errors
        }
    }

    // copy non-formula cell value
    private static void copyCellValue(XSSFCell src, XSSFCell dst) {
        try {
            switch (src.getCellType()) {
                case STRING: dst.setCellValue(src.getRichStringCellValue()); break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(src)) dst.setCellValue(src.getDateCellValue());
                    else dst.setCellValue(src.getNumericCellValue());
                    break;
                case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
                case BLANK: dst.setBlank(); break;
                case ERROR: dst.setCellErrorValue(src.getErrorCellValue()); break;
                default: dst.setCellValue(src.toString()); break;
            }
        } catch (Exception ex) {
            dst.setCellValue(src.toString());
        }
    }

    private static void copyCachedValue(XSSFCell src, XSSFCell dst) {
        try {
            switch (src.getCachedFormulaResultType()) {
                case NUMERIC: dst.setCellValue(src.getNumericCellValue()); break;
                case STRING: dst.setCellValue(src.getStringCellValue()); break;
                case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
                case ERROR: dst.setCellErrorValue(src.getErrorCellValue()); break;
                default: dst.setBlank(); break;
            }
        } catch (Exception ex) {
            dst.setBlank();
        }
    }

    /**
     * Adjusts row numbers in a formula string by delta for references that:
     * - are un-absolute in the row (no $ before row number)
     * - either un-qualified (no sheet) or qualified with source sheet name
     *
     * Additional params srcBlockStart/srcBlockEnd (0-based) are currently unused,
     * but could be used if you want to limit adjustments only for refs inside the block.
     *
     * This is a regex-based approximation and covers common cases: A1, $A$1, A$1, $A1,
     * ranges A1:B2, sheet-qualified 'My Sheet'!A1, Sheet1!A1.
     */
    private static String adjustFormulaRowReferences(String formula, String srcSheetName, int delta, int srcBlockStart0, int srcBlockEnd0) {
        if (delta == 0) return formula;

        List<TextSegment> segs = splitByQuotes(formula);
        StringBuilder out = new StringBuilder();
        for (TextSegment seg : segs) {
            if (seg.inQuotes) {
                out.append(seg.text);
            } else {
                out.append(adjustRefsInUnquotedPart(seg.text, srcSheetName, delta));
            }
        }
        return out.toString();
    }

    // similar helper implementations as in earlier version
    private static String adjustRefsInUnquotedPart(String text, String srcSheetName, int delta) {
        String cellRef = "(?:\\$?[A-Za-z]{1,3}\\$?\\d+)";
        String sheetRef = "(?:'[^']+'|[A-Za-z0-9_\\.]+)!";
        Pattern p = Pattern.compile("(?i)(" + "(?:" + sheetRef + ")?" + cellRef + "(?::" + cellRef + ")?"+ ")");
        Matcher m = p.matcher(text);
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            String token = m.group(1);
            String replaced = adjustSingleRefOrRange(token, srcSheetName, delta);
            m.appendReplacement(sb, Matcher.quoteReplacement(replaced));
        }
        m.appendTail(sb);
        return sb.toString();
    }

    private static String adjustSingleRefOrRange(String token, String srcSheetName, int delta) {
        String sheetQualifier = null;
        String rest = token;
        int excl = token.indexOf('!');
        if (excl >= 0) {
            sheetQualifier = token.substring(0, excl + 1);
            rest = token.substring(excl + 1);
        }

        boolean sheetMatchesSource = false;
        if (sheetQualifier != null) {
            String q = sheetQualifier.substring(0, sheetQualifier.length() - 1);
            String qname = q;
            if (q.startsWith("'") && q.endsWith("'")) qname = q.substring(1, q.length() - 1);
            if (qname.equals(srcSheetName)) sheetMatchesSource = true;
        } else {
            sheetMatchesSource = true;
        }

        if (!sheetMatchesSource) return token;

        if (rest.contains(":")) {
            String[] parts = rest.split(":", 2);
            String left = adjustSingleCellRef(parts[0], delta);
            String right = adjustSingleCellRef(parts[1], delta);
            return left + ":" + right;
        } else {
            return adjustSingleCellRef(rest, delta);
        }
    }

    private static String adjustSingleCellRef(String cellRef, int delta) {
        Pattern p = Pattern.compile("^(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)$");
        Matcher m = p.matcher(cellRef);
        if (!m.find()) return cellRef;
        String colDollar = m.group(1);
        String colLetters = m.group(2);
        String rowDollar = m.group(3);
        String rowNumStr = m.group(4);
        int rowNum = Integer.parseInt(rowNumStr);
        if ("$".equals(rowDollar)) return cellRef; // absolute row -> unchanged
        int newRow = rowNum + delta;
        if (newRow < 1) newRow = 1;
        return colDollar + colLetters + rowDollar + Integer.toString(newRow);
    }

    private static List<TextSegment> splitByQuotes(String formula) {
        List<TextSegment> list = new ArrayList<>();
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < formula.length(); i++) {
            char ch = formula.charAt(i);
            cur.append(ch);
            if (ch == '"') {
                if (i + 1 < formula.length() && formula.charAt(i + 1) == '"') {
                    cur.append('"'); i++;
                } else {
                    list.add(new TextSegment(cur.toString(), inQuotes));
                    cur.setLength(0);
                    inQuotes = !inQuotes;
                }
            }
        }
        if (cur.length() > 0) list.add(new TextSegment(cur.toString(), inQuotes));
        return list;
    }
    private static class TextSegment { String text; boolean inQuotes; TextSegment(String t, boolean q){text=t; inQuotes=q;} }

    // Example usage
    public static void main(String[] args) throws Exception {
        try (FileInputStream fis = new FileInputStream("input.xlsx");
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {

            // copy rows 10..20 from sheet index 0 into sheet index 1 starting at row 50
            copyRowRange(wb, 0, 10, 20, 1, 50);

            try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
                wb.write(fos);
            }
        }
    }
}