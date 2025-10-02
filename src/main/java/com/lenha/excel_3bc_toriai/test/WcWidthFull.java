package com.lenha.excel_3bc_toriai.test;

import java.text.BreakIterator;
import java.util.Locale;

public class WcWidthFull {
    // ----------------------------------------------------------------
    // COMBINING intervals (width = 0)
    // ----------------------------------------------------------------
    private static final int[][] COMBINING = {
            {0x0300, 0x036F}, {0x0483, 0x0489}, {0x0591, 0x05BD}, {0x05BF, 0x05BF},
            {0x05C1, 0x05C2}, {0x05C4, 0x05C5}, {0x05C7, 0x05C7}, {0x0610, 0x061A},
            {0x064B, 0x065F}, {0x0670, 0x0670}, {0x06D6, 0x06DD}, {0x06DF, 0x06E4},
            {0x06E7, 0x06E8}, {0x06EA, 0x06ED}, {0x0711, 0x0711}, {0x0730, 0x074A},
            {0x07A6, 0x07B0}, {0x07EB, 0x07F3}, {0x0816, 0x0819}, {0x081B, 0x0823},
            {0x0825, 0x0827}, {0x0829, 0x082D}, {0x0859, 0x085B}, {0x08D3, 0x08E1},
            {0x08E3, 0x0902}, {0x093A, 0x093A}, {0x093C, 0x093C}, {0x0941, 0x0948},
            {0x094D, 0x094D}, {0x0951, 0x0957}, {0x0962, 0x0963}, {0x0981, 0x0981},
            {0x09BC, 0x09BC}, {0x09C1, 0x09C4}, {0x09CD, 0x09CD}, {0x09E2, 0x09E3},
            {0x0A01, 0x0A02}, {0x0A3C, 0x0A3C}, {0x0A41, 0x0A42}, {0x0A47, 0x0A48},
            {0x0A4B, 0x0A4D}, {0x0A51, 0x0A51}, {0x0A70, 0x0A71}, {0x0A75, 0x0A75},
            {0x0A81, 0x0A82}, {0x0ABC, 0x0ABC}, {0x0AC1, 0x0AC5}, {0x0AC7, 0x0AC8},
            {0x0ACD, 0x0ACD}, {0x0AE2, 0x0AE3}, {0x0B01, 0x0B01}, {0x0B3C, 0x0B3C},
            {0x0B3F, 0x0B3F}, {0x0B41, 0x0B43}, {0x0B4D, 0x0B4D}, {0x0B56, 0x0B56},
            {0x0B82, 0x0B82}, {0x0BC0, 0x0BC0}, {0x0BCD, 0x0BCD}, {0x0C3E, 0x0C40},
            {0x0C46, 0x0C48}, {0x0C4A, 0x0C4D}, {0x0C55, 0x0C56}, {0x0CBC, 0x0CBC},
            {0x0CBF, 0x0CBF}, {0x0CC6, 0x0CC6}, {0x0CCC, 0x0CCD}, {0x0CE2, 0x0CE3},
            {0x0D41, 0x0D43}, {0x0D4D, 0x0D4D}, {0x0DCA, 0x0DCA}, {0x0DD2, 0x0DD4},
            {0x0DD6, 0x0DD6}, {0x0E31, 0x0E31}, {0x0E34, 0x0E3A}, {0x0E47, 0x0E4E},
            {0x0EB1, 0x0EB1}, {0x0EB4, 0x0EB9}, {0x0EBB, 0x0EBC}, {0x0EC8, 0x0ECD},
            {0x0F18, 0x0F19}, {0x0F35, 0x0F35}, {0x0F37, 0x0F37}, {0x0F39, 0x0F39},
            {0x0F71, 0x0F7E}, {0x0F80, 0x0F84}, {0x0F86, 0x0F87}, {0x0F90, 0x0F97},
            {0x0F99, 0x0FBC}, {0x0FC6, 0x0FC6}, {0x102D, 0x1030}, {0x1032, 0x1032},
            {0x1036, 0x1037}, {0x1039, 0x1039}, {0x1058, 0x1059}, {0x1160, 0x11FF},
            {0x135F, 0x135F}, {0x1712, 0x1714}, {0x1732, 0x1734}, {0x1752, 0x1753},
            {0x1772, 0x1773}, {0x17B4, 0x17B5}, {0x17B7, 0x17BD}, {0x17C6, 0x17C6},
            {0x17C9, 0x17D3}, {0x17DD, 0x17DD}, {0x180B, 0x180D}, {0x18A9, 0x18A9},
            {0x1920, 0x1922}, {0x1927, 0x1928}, {0x1932, 0x1932}, {0x1939, 0x193B},
            {0x1A17, 0x1A18}, {0x1B00, 0x1B03}, {0x1B34, 0x1B34}, {0x1B36, 0x1B3A},
            {0x1B3C, 0x1B3C}, {0x1B42, 0x1B42}, {0x1B6B, 0x1B73}, {0x1DC0, 0x1DCA},
            {0x1DFE, 0x1DFF}, {0x200B, 0x200F}, {0x202A, 0x202E}, {0x2060, 0x2063},
            {0x206A, 0x206F}, {0x20D0, 0x20EF}, {0x302A, 0x302F}, {0x3099, 0x309A},
            {0xA66F, 0xA66F}, {0xA670, 0xA672}, {0xA674, 0xA67D}, {0xA69E, 0xA69F},
            {0xA6F0, 0xA6F1}, {0xA802, 0xA802}, {0xA806, 0xA806}, {0xA80B, 0xA80B},
            {0xA823, 0xA827}, {0xA880, 0xA881}, {0xA8B4, 0xA8C5}, {0xA8E0, 0xA8F1},
            {0xA926, 0xA92D}, {0xA947, 0xA951}, {0xA980, 0xA982}, {0xA9B3, 0xA9B3},
            {0xA9B6, 0xA9B9}, {0xA9BC, 0xA9BD}, {0xA9E5, 0xA9E5}, {0xAA29, 0xAA2E},
            {0xAA31, 0xAA32}, {0xAA35, 0xAA36}, {0xAA43, 0xAA43}, {0xAA4C, 0xAA4C},
            {0xAA7C, 0xAA7C}, {0xAAB0, 0xAAB0}, {0xAAB2, 0xAAB4}, {0xAAB7, 0xAAB8},
            {0xAABE, 0xAABF}, {0xAAC1, 0xAAC1}, {0xAAEB, 0xAAEF}, {0xAAF5, 0xAAF6},
            {0xABE3, 0xABE4}, {0xABE6, 0xABE7}, {0xABE9, 0xABEA}, {0xABEC, 0xABEC},
            {0x11000, 0x11002}, {0x11038, 0x11046}, {0x1107F, 0x11082}, {0x110B0, 0x110BA},
            {0x11100, 0x11102}, {0x11127, 0x1112B}, {0x1112D, 0x11134}, {0x11173, 0x11173},
            {0x11180, 0x11182}, {0x111B3, 0x111C0}, {0x111C9, 0x111CC}, {0x1122C, 0x11237},
            {0x1123E, 0x1123E}, {0x112DF, 0x112EA}, {0x11300, 0x11301}, {0x1133B, 0x1133C},
            {0x11340, 0x11340}, {0x11366, 0x1136C}, {0x11370, 0x11374}, {0x11438, 0x1143F},
            {0x11442, 0x11444}, {0x11446, 0x11446}, {0x1145E, 0x1145E}, {0x114B3, 0x114B8},
            {0x114BA, 0x114BA}, {0x114BF, 0x114C0}, {0x114C2, 0x114C3}, {0x115B2, 0x115B5},
            {0x115BC, 0x115BD}, {0x115BF, 0x115C0}, {0x115DC, 0x115DD}, {0x11633, 0x1163A},
            {0x1163D, 0x1163D}, {0x1163F, 0x11640}, {0x116AB, 0x116AB}, {0x116AD, 0x116AD},
            {0x116B0, 0x116B5}, {0x116B7, 0x116B7}, {0x1171D, 0x1171F}, {0x11722, 0x11725},
            {0x11727, 0x1172B}, {0x1182F, 0x11837}, {0x11839, 0x1183A}, {0x1193B, 0x1193C},
            {0x1193E, 0x1193E}, {0x11943, 0x11943}, {0x119D4, 0x119D7}, {0x119DA, 0x119DB},
            {0x119E0, 0x119E0}, {0x11A01, 0x11A0A}, {0x11A33, 0x11A38}, {0x11A3B, 0x11A3E},
            {0x11A47, 0x11A47}, {0x11A51, 0x11A5B}, {0x11A8A, 0x11A96}, {0x11A98, 0x11A99},
            {0x11C30, 0x11C36}, {0x11C38, 0x11C3D}, {0x11C3F, 0x11C3F}, {0x11C92, 0x11CA7},
            {0x11CAA, 0x11CB0}, {0x11CB2, 0x11CB3}, {0x11CB5, 0x11CB6}, {0x11D31, 0x11D36},
            {0x11D3A, 0x11D3A}, {0x11D3C, 0x11D3D}, {0x11D3F, 0x11D45}, {0x11D47, 0x11D47},
            {0x11D90, 0x11D91}, {0x11D95, 0x11D95}, {0x11D97, 0x11D97}, {0x11EF3, 0x11EF4},
            {0x16AF0, 0x16AF4}, {0x16B30, 0x16B36}, {0x16F8F, 0x16F92}, {0x1BC9D, 0x1BC9E},
            {0x1BCA0, 0x1BCA3}, {0x1D167, 0x1D169}, {0x1D17B, 0x1D182}, {0x1D185, 0x1D18B},
            {0x1D1AA, 0x1D1AD}, {0xE0100, 0xE01EF}
    };

    // ----------------------------------------------------------------
    // WIDE intervals (width = 2)
    // ----------------------------------------------------------------
    private static final int[][] WIDE = {
            {0x1100, 0x115F}, {0x2329, 0x232A}, {0x2E80, 0xA4CF},
            {0xAC00, 0xD7A3}, {0xF900, 0xFAFF}, {0xFE10, 0xFE19},
            {0xFE30, 0xFE6F}, {0xFF00, 0xFF60}, {0xFFE0, 0xFFE6},
            {0x1F300, 0x1F64F}, {0x1F900, 0x1F9FF}, {0x20000, 0x2FFFD},
            {0x30000, 0x3FFFD}
    };

    // ----------------------------------------------------------------
    // AMBIGUOUS intervals (treat as 1 or 2 depending on policy)
    // ----------------------------------------------------------------
    private static final int[][] AMBIGUOUS = {
            {0x00A1, 0x00A1}, {0x00A4, 0x00A4}, {0x00A7, 0x00A8}, {0x00AA, 0x00AA},
            {0x00AD, 0x00AE}, {0x00B0, 0x00B4}, {0x00B6, 0x00BA}, {0x00BC, 0x00BF},
            {0x00C6, 0x00C6}, {0x00D0, 0x00D0}, {0x00D7, 0x00D8}, {0x00DE, 0x00E1},
            {0x00E6, 0x00E6}, {0x00E8, 0x00EA}, {0x00EC, 0x00ED}, {0x00F0, 0x00F0},
            {0x00F2, 0x00F3}, {0x00F7, 0x00FA}, {0x00FC, 0x00FC}, {0x00FE, 0x00FE},
            {0x0101, 0x0101}, {0x0111, 0x0111}, {0x0113, 0x0113}, {0x011B, 0x011B},
            {0x0126, 0x0127}, {0x012B, 0x012B}, {0x0131, 0x0133}, {0x0138, 0x0138},
            {0x013F, 0x0142}, {0x0144, 0x0144}, {0x0148, 0x014B}, {0x014D, 0x014D},
            {0x0152, 0x0153}, {0x0166, 0x0167}, {0x016B, 0x016B}, {0x01CE, 0x01CE},
            {0x01D0, 0x01D0}, {0x01D2, 0x01D2}, {0x01D4, 0x01D4}, {0x01D6, 0x01D6},
            {0x01D8, 0x01D8}, {0x01DA, 0x01DA}, {0x01DC, 0x01DC}, {0x0251, 0x0251},
            {0x0261, 0x0261}, {0x02C4, 0x02C4}, {0x02C7, 0x02C7}, {0x02C9, 0x02CB},
            {0x02CD, 0x02CD}, {0x02D0, 0x02D0}, {0x02D8, 0x02DB}, {0x02DD, 0x02DD},
            {0x02DF, 0x02DF}, {0x0300, 0x036F}, {0x0391, 0x03A1}, {0x03A3, 0x03A9},
            {0x03B1, 0x03C1}, {0x03C3, 0x03C9}, {0x0401, 0x0401}, {0x0410, 0x044F},
            {0x0451, 0x0451}, {0x2010, 0x2010}, {0x2013, 0x2016}, {0x2018, 0x2019},
            {0x201C, 0x201D}, {0x2020, 0x2022}, {0x2026, 0x2026}, {0x2030, 0x2030},
            {0x2032, 0x2033}, {0x2035, 0x2035}, {0x203B, 0x203B}, {0x203E, 0x203E},
            {0x2074, 0x2074}, {0x207F, 0x207F}, {0x2081, 0x2084}, {0x20AC, 0x20AC},
            {0x2103, 0x2103}, {0x2105, 0x2105}, {0x2109, 0x2109}, {0x2113, 0x2113},
            {0x2116, 0x2116}, {0x2121, 0x2122}, {0x2126, 0x2126}, {0x212B, 0x212B},
            {0x2153, 0x2154}, {0x215B, 0x215E}, {0x2160, 0x216B}, {0x2170, 0x2179},
            {0x2190, 0x2199}, {0x21B8, 0x21B9}, {0x21D2, 0x21D2}, {0x21D4, 0x21D4},
            {0x21E7, 0x21E7}, {0x2200, 0x2200}, {0x2202, 0x2203}, {0x2207, 0x2208},
            {0x220B, 0x220B}, {0x220F, 0x220F}, {0x2211, 0x2211}, {0x2215, 0x2215},
            {0x221A, 0x221A}, {0x221D, 0x2220}, {0x2223, 0x2223}, {0x2225, 0x2225},
            {0x2227, 0x222C}, {0x222E, 0x222E}, {0x2234, 0x2237}, {0x223C, 0x223D},
            {0x2248, 0x2248}, {0x224C, 0x224C}, {0x2252, 0x2252}, {0x2260, 0x2261},
            {0x2264, 0x2267}, {0x226A, 0x226B}, {0x226E, 0x226F}, {0x2282, 0x2283},
            {0x2286, 0x2287}, {0x2295, 0x2295}, {0x2299, 0x2299}, {0x22A5, 0x22A5},
            {0x22BF, 0x22BF}, {0x2312, 0x2312}, {0x2460, 0x24E9}, {0x24EB, 0x254B},
            {0x2550, 0x2573}, {0x2580, 0x258F}, {0x2592, 0x2595}, {0x25A0, 0x25A1},
            {0x25A3, 0x25A9}, {0x25B2, 0x25B3}, {0x25B6, 0x25B7}, {0x25BC, 0x25BD},
            {0x25C0, 0x25C1}, {0x25C6, 0x25C8}, {0x25CB, 0x25CB}, {0x25CE, 0x25D1},
            {0x25E2, 0x25E5}, {0x25EF, 0x25EF}, {0x2605, 0x2606}, {0x2609, 0x2609},
            {0x260E, 0x260F}, {0x2614, 0x2615}, {0x261C, 0x261C}, {0x261E, 0x261E},
            {0x2640, 0x2640}, {0x2642, 0x2642}, {0x2660, 0x2661}, {0x2663, 0x2665},
            {0x2667, 0x266A}, {0x266C, 0x266D}, {0x266F, 0x266F}, {0x269E, 0x269F},
            {0x26BF, 0x26BF}, {0x26C6, 0x26CD}, {0x26CF, 0x26D3}, {0x26D5, 0x26E1},
            {0x26E3, 0x26E3}, {0x26E8, 0x26E9}, {0x26EB, 0x26F1}, {0x26F4, 0x26F4},
            {0x26F6, 0x26F9}, {0x26FB, 0x26FC}, {0x26FE, 0x26FF}, {0x273D, 0x273D},
            {0x2776, 0x277F}, {0x2B56, 0x2B59}, {0x3248, 0x324F}, {0xE000, 0xF8FF},
            {0xFE00, 0xFE0F}, {0xFFFD, 0xFFFD}
    };

    private WcWidthFull() { /* utility */ }

    // ----------------------------------------------------------------
    // Binary search helper for interval tables
    // ----------------------------------------------------------------
    private static boolean inIntervals(int[][] table, int codePoint) {
        int lo = 0, hi = table.length - 1;
        while (lo <= hi) {
            int mid = (lo + hi) >>> 1;
            int[] iv = table[mid];
            if (codePoint < iv[0]) hi = mid - 1;
            else if (codePoint > iv[1]) lo = mid + 1;
            else return true;
        }
        return false;
    }

    // ----------------------------------------------------------------
    // wcwidth for a single code point
    // returns:
    //  -1 : control chars (C0/C1)
    //   0 : combining marks
    //   1 : narrow
    //   2 : wide
    // treatAmbiguousAsWide: if true, AMBIGUOUS intervals treated as wide (2)
    // ----------------------------------------------------------------
    public static int wcwidth(int codePoint, boolean treatAmbiguousAsWide) {
        if (codePoint == 0) return 0;
        if (codePoint < 32 || (codePoint >= 0x7f && codePoint < 0xa0)) return -1;
        if (inIntervals(COMBINING, codePoint)) return 0;
        if (inIntervals(WIDE, codePoint)) return 2;
        if (inIntervals(AMBIGUOUS, codePoint)) return treatAmbiguousAsWide ? 2 : 1;
        return 1;
    }

    public static int wcwidth(int codePoint) {
        return wcwidth(codePoint, false);
    }

    // ----------------------------------------------------------------
    // wcswidth: compute width of string by summing codepoint widths
    // returns -1 if any control char encountered
    // ----------------------------------------------------------------
    public static int wcswidth(String s, boolean treatAmbiguousAsWide) {
        if (s == null || s.isEmpty()) return 0;
        int width = 0;
        int i = 0, len = s.length();
        while (i < len) {
            int cp = s.codePointAt(i);
            int w = wcwidth(cp, treatAmbiguousAsWide);
            if (w < 0) return -1;
            width += w;
            i += Character.charCount(cp);
        }
        return width;
    }

    public static int wcswidth(String s) {
        return wcswidth(s, false);
    }

    // ----------------------------------------------------------------
    // Grapheme-aware cluster width
    // - Treat a whole grapheme cluster as wide (2) if:
    //   * it contains a ZWJ (U+200D) OR
    //   * it contains VS16 (U+FE0F) OR
    //   * any code point inside is wide (wcwidth == 2) OR
    //   * it is exactly two regional indicators (flag)
    // - Otherwise sum per-codepoint widths (combining marks -> 0)
    // Returns -1 on control char found.
    // ----------------------------------------------------------------
    private static int clusterWidth(String cluster, boolean treatAmbiguousAsWide) {
        if (cluster == null || cluster.isEmpty()) return 0;

        boolean hasZWJ = false;
        boolean hasVS16 = false;
        boolean anyWideCodepoint = false;
        int regionalIndicatorCount = 0;

        int i = 0, len = cluster.length();
        while (i < len) {
            int cp = cluster.codePointAt(i);
            // control
            if (cp == 0) return 0;
            if (cp < 32 || (cp >= 0x7f && cp < 0xa0)) return -1;
            if (cp == 0x200D) hasZWJ = true;
            if (cp == 0xFE0F) hasVS16 = true;
            if (cp >= 0x1F1E6 && cp <= 0x1F1FF) regionalIndicatorCount++;
            int w = wcwidth(cp, treatAmbiguousAsWide);
            if (w == 2) anyWideCodepoint = true;
            i += Character.charCount(cp);
        }

        if (hasZWJ || hasVS16 || anyWideCodepoint) {
            return 2;
        }

        if (regionalIndicatorCount == 2) {
            return 2;
        }

        // fallback: sum per-codepoint widths
        int sum = 0;
        i = 0;
        while (i < len) {
            int cp = cluster.codePointAt(i);
            int w = wcwidth(cp, treatAmbiguousAsWide);
            if (w < 0) return -1;
            sum += w;
            i += Character.charCount(cp);
        }
        return sum;
    }

    // ----------------------------------------------------------------
    // Grapheme-aware wcswidth: split using BreakIterator.getCharacterInstance()
    // (which approximates extended grapheme clusters), compute cluster widths.
    // Returns -1 if any control char found.
    // ----------------------------------------------------------------
    public static int wcswidthGrapheme(String s, boolean treatAmbiguousAsWide) {
        if (s == null || s.isEmpty()) return 0;
        BreakIterator it = BreakIterator.getCharacterInstance(Locale.ROOT);
        it.setText(s);
        int start = it.first();
        int width = 0;
        for (int end = it.next(); end != BreakIterator.DONE; start = end, end = it.next()) {
            String cluster = s.substring(start, end);
            int cw = clusterWidth(cluster, treatAmbiguousAsWide);
            if (cw < 0) return -1;
            width += cw;
        }
        return width;
    }

    public static int wcswidthGrapheme(String s) {
        return wcswidthGrapheme(s, false);
    }

    // ----------------------------------------------------------------
    // Demo main

    // ----------------------------------------------------------------

    /**
     * Ghi chú ngắn:
     *
     * treatAmbiguousAsWide = true hữu ích khi chạy trong môi trường CJK/terminal mà các ký tự ambiguous được hiển thị như wide (2). Mặc định mình để false (ambiguous = 1).
     *
     * wcswidthGrapheme dùng BreakIterator.getCharacterInstance(Locale.ROOT) để tách grapheme clusters — phù hợp với đa số trường hợp (base+combining, ZWJ sequences, flags). Nếu bạn muốn độ chính xác cao hơn cho edge-case emoji, ta có thể thay bằng ICU4J (grapheme cluster iterator theo UAX#29/Unicode Emoji rules).
     *
     * Hàm trả -1 nếu gặp control characters (C0/C1). Nếu muốn bỏ qua control thì thay xử lý thành 0 thay vì -1.
     *
     * Nếu bạn muốn:
     *
     * mình đóng gói thêm unit tests (JUnit 5) cho các trường hợp cạnh (combining, ZWJ, flags, ambiguous), hoặc
     *
     * chuyển BreakIterator -> ICU4J để xử lý grapheme clusters theo Unicode Emoji spec,
     *
     * nói “unit tests” hoặc “ICU4J” — mình sẽ dán code tiếp.
     * @param args
     */
    public static void main(String[] args) {
        String[] tests = {
                "a",
                "Z",
                "あ",               // hiragana
                "漢",               // kanji
                "e\u0301",          // e + combining acute -> should be 1
                "👍",               // thumbs up emoji -> 2
                "👨‍👩‍👧‍👦",       // family ZWJ sequence -> 2
                "🇯🇵",              // flag (regional indicators) -> 2
                "aあb漢c",
                "·",                // middle dot (ambiguous)
                "—"                 // em dash (ambiguous)
        };

        System.out.println("=== wcswidth (codepoint-sum) ===");
        for (String t : tests) {
            System.out.printf("%-10s -> narrow=%d, wide=%d%n",
                    t,
                    wcswidth(t, false),
                    wcswidth(t, true));
        }

        System.out.println("\n=== wcswidthGrapheme (grapheme clusters) ===");
        for (String t : tests) {
            System.out.printf("%-10s -> narrow=%d, wide=%d%n",
                    t,
                    wcswidthGrapheme(t, false),
                    wcswidthGrapheme(t, true));
        }
    }
}
