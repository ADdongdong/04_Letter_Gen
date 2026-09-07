package com.xidian.lettergen;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHeightRule;

/**
 * 制函渲染引擎（Java/POI 移植，功能对齐 Python 版 backend/python/render_engine.py）：
 * - readExcelSheets / readExcelGrouped：Excel 读取与批量分组（前两列「被审计单位/被询证单位」）
 * - extractPlaceholderBindings / renderTemplate：占位段驱动的制函（模板即定义）
 * - injectSheetTable：表格注入（样式配置：字号/字体 eastAsia/对齐/行高 atLeast/列宽三档）
 * - 占位段精确匹配 + 终扫兜底，无数据/缺失 Sheet 占位段一律删除
 */
public final class RenderEngine {

    private RenderEngine() {}

    // ===== 表格样式默认值（与 Python DEFAULT_TABLE_STYLE 对齐，缺省逐字段回落=原硬编码行为）=====
    public static Map<String, Object> defaultTableStyle() {
        Map<String, Object> s = new LinkedHashMap<>();
        s.put("font_size", 10);
        s.put("font_name", "");
        s.put("h_align", "center");
        s.put("v_align", "top");
        s.put("row_height_mode", "auto");
        s.put("row_height_pt", 20);
        s.put("col_widths_mode", "equal");
        s.put("col_widths", new LinkedHashMap<String, Object>());
        s.put("col_widths_ratio", "");
        return s;
    }

    /** 用户样式与默认值合并：逐字段类型/枚举校验，非法值回落默认（防脏数据） */
    @SuppressWarnings("unchecked")
    public static Map<String, Object> resolveStyle(Map<String, Object> style) {
        Map<String, Object> s = defaultTableStyle();
        if (style == null) return s;
        Object v;
        v = style.get("font_size");
        if (v instanceof Number n && n.doubleValue() >= 5 && n.doubleValue() <= 72) s.put("font_size", n.doubleValue());
        v = style.get("font_name");
        if (v instanceof String t) s.put("font_name", t.trim());
        v = style.get("h_align");
        if ("left".equals(v) || "center".equals(v) || "right".equals(v)) s.put("h_align", v);
        v = style.get("v_align");
        if ("top".equals(v) || "center".equals(v)) s.put("v_align", v);
        v = style.get("row_height_mode");
        if ("auto".equals(v) || "atLeast".equals(v)) s.put("row_height_mode", v);
        v = style.get("row_height_pt");
        if (v instanceof Number n && n.doubleValue() >= 5 && n.doubleValue() <= 200) s.put("row_height_pt", n.doubleValue());
        v = style.get("col_widths_mode");
        if ("equal".equals(v) || "auto".equals(v) || "ratio".equals(v)) s.put("col_widths_mode", v);
        v = style.get("col_widths");
        if (v instanceof Map<?, ?> m) {
            Map<String, Object> cw = new LinkedHashMap<>();
            m.forEach((k, val) -> cw.put(String.valueOf(k), val == null ? "" : String.valueOf(val)));
            s.put("col_widths", cw);
        }
        v = style.get("col_widths_ratio");
        if (v instanceof String t) s.put("col_widths_ratio", t.trim());
        return s;
    }

    /** '2:1:1' -> [2.0,1.0,1.0]；空/非法/列数不匹配返回 null（回落均分） */
    public static List<Double> parseColRatio(String ratioStr, int ncols) {
        if (ratioStr == null || ratioStr.isEmpty()) return null;
        String[] parts = ratioStr.split(":");
        List<Double> out = new ArrayList<>();
        try {
            for (String p : parts) if (!p.isBlank()) out.add(Double.parseDouble(p.trim()));
        } catch (NumberFormatException e) {
            return null;
        }
        double sum = out.stream().mapToDouble(Double::doubleValue).sum();
        if (out.size() != ncols || sum <= 0 || out.stream().anyMatch(d -> d < 0)) return null;
        return out;
    }

    /** 单字符显示宽度：CJK/全角按 2，其余按 1（对齐 Python unicodedata.east_asian_width W/F） */
    private static int charW(char ch) {
        int t = Character.getType(ch);
        return (t == Character.OTHER_LETTER || t == Character.OTHER_SYMBOL
                || t == Character.NON_SPACING_MARK || t == Character.ENCLOSING_MARK) ? 2 : 1;
    }

    private static int strW(String s) {
        int w = 0;
        for (char c : s.toCharArray()) w += charW(c);
        return w;
    }

    /** 按内容自适应列宽权重：每列 = max(表头宽, 该列数据最大宽)，下限 1；列数 0 返回 null */
    public static List<Double> colWeightsAuto(SheetData sheet) {
        List<String> header = sheet.header;
        int ncols = header.size();
        if (ncols == 0) return null;
        List<Double> weights = new ArrayList<>();
        for (String h : header) weights.add((double) strW(h));
        for (List<String> row : sheet.rows) {
            for (int j = 0; j < ncols; j++) {
                int w = strW(j < row.size() ? row.get(j) : "");
                if (w > weights.get(j)) weights.set(j, (double) w);
            }
        }
        return weights.stream().map(w -> Math.max(w, 1.0)).toList();
    }

    /** 单个 Sheet 的内存结构（对齐 Python 的 sheet dict） */
    public static class SheetData {
        public String name = "";
        public List<String> header = new ArrayList<>();
        public List<List<String>> rows = new ArrayList<>();
        public List<Integer> rowNums = new ArrayList<>();
    }

    /** read_excel_grouped 的返回结构（对齐 Python dict 键） */
    public static class GroupedResult {
        public List<String> sheets = new ArrayList<>();
        public List<String[]> groups = new ArrayList<>();          // [audit, confirm]
        public List<Map<String, Object>> emptyCells = new ArrayList<>();  // {sheet,row,col}
        public Map<String, Map<String, SheetData>> virtualSheets = new LinkedHashMap<>();  // sheetName -> (audit+"\n"+confirm -> virtual)
    }

    // ===== 占位段常量 =====
    public static final String PLACEHOLDER_PREFIX = "【Sheet";

    public static String placeholderText(String sheetName) {
        return "【Sheet「" + sheetName + "」表格将在此处展示】";
    }

    public static boolean isGroupHeader(List<String> header) {
        if (header == null || header.size() < 2) return false;
        return "被审计单位".equals(header.get(0).trim()) && "被询证单位".equals(header.get(1).trim());
    }

    /** 读取 Excel 所有 Sheet：表头=第一行；过滤全空数据行，保留 Excel 真实行号（1 基） */
    public static List<SheetData> readExcelSheets(String xlsxPath) throws IOException {
        List<SheetData> out = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(xlsxPath); XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                XSSFSheet ws = wb.getSheetAt(i);
                SheetData sd = new SheetData();
                sd.name = wb.getSheetName(i);
                int lastRow = ws.getLastRowNum();
                if (ws.getPhysicalNumberOfRows() == 0 || lastRow < 0) continue;
                XSSFRow headRow = ws.getRow(0);
                if (headRow == null) continue;
                int maxCol = 0;
                for (int c = 0; c <= headRow.getLastCellNum(); c++) {
                    XSSFCell cell = headRow.getCell(c);
                    if (cell != null) maxCol = Math.max(maxCol, cell.getColumnIndex() + 1);
                }
                if (maxCol == 0) continue;
                for (int c = 0; c < maxCol; c++) sd.header.add(cellStr(headRow.getCell(c)));
                for (int r = 1; r <= lastRow; r++) {
                    XSSFRow row = ws.getRow(r);
                    List<String> cells = new ArrayList<>();
                    boolean hasContent = false;
                    for (int c = 0; c < maxCol; c++) {
                        String v = row == null ? "" : cellStr(row.getCell(c));
                        if (!v.isBlank()) hasContent = true;
                        cells.add(v);
                    }
                    if (hasContent) {
                        sd.rows.add(cells);
                        sd.rowNums.add(r + 1);  // 1 基 Excel 行号
                    }
                }
                out.add(sd);
            }
        }
        return out;
    }

    /** 单元格转文本，数值格式对齐 Python str()：整数值无小数尾巴 */
    private static String cellStr(XSSFCell cell) {
        if (cell == null) return "";
        String s = switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double d = cell.getNumericCellValue();
                yield (d == Math.floor(d) && !Double.isInfinite(d) && Math.abs(d) < 1e15)
                        ? String.valueOf((long) d) : String.valueOf(d);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                // data_only 语义：公式取缓存计算值
                try { yield String.valueOf(cell.getNumericCellValue()); }
                catch (Exception e) { yield cell.getStringCellValue(); }
            }
            default -> "";
        };
        return s == null ? "" : s;
    }

    /** 批量 Excel 分组读取：仅处理表头前两列为「被审计单位/被询证单位」的 Sheet */
    public static GroupedResult readExcelGrouped(String xlsxPath) throws IOException {
        List<SheetData> sheets = readExcelSheets(xlsxPath);
        GroupedResult res = new GroupedResult();
        Map<String, Map<String, List<List<String>>>> groupedRows = new LinkedHashMap<>();
        Map<String, List<String>> sheetHeaders = new LinkedHashMap<>();
        java.util.Set<String> seen = new java.util.HashSet<>();
        for (SheetData sheet : sheets) {
            if (!isGroupHeader(sheet.header)) continue;
            Map<String, List<List<String>>> buckets = new LinkedHashMap<>();
            for (int i = 0; i < sheet.rows.size(); i++) {
                List<String> row = sheet.rows.get(i);
                String audit = row.size() > 0 ? row.get(0).trim() : "";
                String confirm = row.size() > 1 ? row.get(1).trim() : "";
                if (audit.isEmpty()) {
                    Map<String, Object> ec = new LinkedHashMap<>();
                    ec.put("sheet", sheet.name); ec.put("row", sheet.rowNums.get(i)); ec.put("col", "被审计单位");
                    res.emptyCells.add(ec);
                }
                if (confirm.isEmpty()) {
                    Map<String, Object> ec = new LinkedHashMap<>();
                    ec.put("sheet", sheet.name); ec.put("row", sheet.rowNums.get(i)); ec.put("col", "被询证单位");
                    res.emptyCells.add(ec);
                }
                String key = audit + "\n" + confirm;
                buckets.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
                if (seen.add(key)) res.groups.add(new String[]{audit, confirm});
            }
            groupedRows.put(sheet.name, buckets);
            sheetHeaders.put(sheet.name, new ArrayList<>(sheet.header.subList(2, sheet.header.size())));
            res.sheets.add(sheet.name);
        }
        // 为每个分组 Sheet × 每个全局组合构造虚拟 sheet（无数据的组合 rows=[]）
        for (String name : res.sheets) {
            Map<String, SheetData> vmap = new LinkedHashMap<>();
            for (String[] key : res.groups) {
                SheetData vs = new SheetData();
                vs.name = name;
                vs.header = sheetHeaders.get(name);
                String k = key[0] + "\n" + key[1];
                for (List<String> r : groupedRows.get(name).getOrDefault(k, List.of())) {
                    vs.rows.add(new ArrayList<>(r.subList(2, r.size())));
                }
                vmap.put(k, vs);
            }
            res.virtualSheets.put(name, vmap);
        }
        return res;
    }

    // ===== docx 底层辅助 =====

    /** XMLBean 属性值（union 类型可能为 Number 或 XmlObject）转 long */
    private static long toLong(Object v) {
        if (v instanceof Number n) return n.longValue();
        if (v instanceof org.apache.xmlbeans.SimpleValue sv) return Long.parseLong(sv.getStringValue().trim());
        return Long.parseLong(String.valueOf(v).trim());
    }

    /** 页面文本宽度（twips）：w:pgSz@w - 左右页边距 */
    public static int textWidthTwips(XWPFDocument doc) {
        var sectPr = doc.getDocument().getBody().getSectPr();
        long pageW = 11906, left = 1800, right = 1800;  // 默认 A4/常规边距
        if (sectPr != null) {
            if (sectPr.isSetPgSz() && sectPr.getPgSz().isSetW()) pageW = toLong(sectPr.getPgSz().getW());
            if (sectPr.isSetPgMar()) {
                var mar = sectPr.getPgMar();
                left = toLong(mar.getLeft());
                right = toLong(mar.getRight());
            }
        }
        return (int) Math.max(pageW - left - right, 1000);
    }

    /** 单元格文本+样式：水平对齐/字号/加粗/字体（setFontFamily 同时设置 eastAsia）/垂直对齐 */
    private static void setCellText(XWPFTableCell cell, String text, boolean bold, double size,
                                    String align, String fontName, String vAlign) {
        cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.setAlignment(switch (align) {
            case "left" -> ParagraphAlignment.LEFT;
            case "right" -> ParagraphAlignment.RIGHT;
            default -> ParagraphAlignment.CENTER;
        });
        XWPFRun run = p.createRun();
        run.setText(text);
        run.setBold(bold);
        run.setFontSize((int) Math.round(size));
        if (fontName != null && !fontName.isEmpty()) run.setFontFamily(fontName);
        if ("center".equals(vAlign)) cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    }

    /** 行最小高度（trHeight hRule=atLeast：内容少保持设定值，内容多自动加高不截断） */
    private static void setRowMinHeight(XWPFTableRow row, double pt) {
        CTTrPr trPr = row.getCtRow().isSetTrPr() ? row.getCtRow().getTrPr() : row.getCtRow().addNewTrPr();
        var trHeight = trPr.sizeOfTrHeightArray() > 0 ? trPr.getTrHeightArray(0) : trPr.addNewTrHeight();
        trHeight.setVal(BigInteger.valueOf((long) (pt * 20)));  // 磅 -> twips
        trHeight.setHRule(STHeightRule.AT_LEAST);
    }

    /** 表格边框：直接写 tblBorders（6 向单线），不依赖模板 styles.xml 是否有 TableGrid 样式 */
    private static void applyGridBorders(XWPFTable table) {
        CTTblPr pr = table.getCTTbl().getTblPr() != null ? table.getCTTbl().getTblPr() : table.getCTTbl().addNewTblPr();
        var borders = pr.isSetTblBorders() ? pr.getTblBorders() : pr.addNewTblBorders();
        for (var b : new java.util.HashMap<String, org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.Enum>() {{
            put("top", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
            put("bottom", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
            put("left", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
            put("right", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
            put("insideH", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
            put("insideV", org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.SINGLE);
        }}.entrySet()) {
            var border = switch (b.getKey()) {
                case "top" -> borders.isSetTop() ? borders.getTop() : borders.addNewTop();
                case "bottom" -> borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom();
                case "left" -> borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft();
                case "right" -> borders.isSetRight() ? borders.getRight() : borders.addNewRight();
                case "insideH" -> borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH();
                default -> borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV();
            };
            border.setVal(b.getValue());
            border.setSz(BigInteger.valueOf(4));
            border.setColor("auto");
        }
    }

    /** 把表格移动到指定段落之后（占位段位置注入）：
     *  toEndToken 定位闭合标签 → toNextToken 跳到 body 层级（段落之后）→ moveXml 插入到该位置之前 */
    public static void moveTableAfterParagraph(XWPFDocument doc, XWPFTable table, int paraIdx) {
        List<XWPFParagraph> paras = doc.getParagraphs();
        if (paraIdx < 0 || paraIdx >= paras.size()) return;
        try (var anchorCursor = paras.get(paraIdx).getCTP().newCursor();
             var tblCursor = table.getCTTbl().newCursor()) {
            anchorCursor.toEndToken();
            anchorCursor.toNextToken();
            tblCursor.moveXml(anchorCursor);
        }
    }

    /** 删除指定索引的 body 顶层段落（POI 官方 removeBodyElement：同步内部 paragraphs 缓存列表） */
    public static void removeParagraphAt(XWPFDocument doc, int idx) {
        List<XWPFParagraph> paras = doc.getParagraphs();
        if (idx < 0 || idx >= paras.size()) return;
        XWPFParagraph para = paras.get(idx);
        int pos = doc.getPosOfParagraph(para);
        if (pos >= 0) doc.removeBodyElement(pos);
    }

    /** 查找占位段（occurrence：同一 Sheet 多处插入时取第 N 次出现，0 基），返回索引或 -1 */
    public static int findPlaceholderPara(XWPFDocument doc, String sheetName, int occurrence) {
        String marker = "Sheet「" + sheetName + "」";
        int seen = -1;
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (int i = 0; i < paras.size(); i++) {
            String t = paras.get(i).getText();
            if (t != null && t.contains(PLACEHOLDER_PREFIX) && t.contains(marker)) {
                seen++;
                if (seen >= occurrence) return i;
            }
        }
        return -1;
    }

    /** 扫描模板中所有占位段自动生成 bindings（模板即定义），同一 Sheet 多处生成多条 */
    public static List<Map<String, Object>> extractPlaceholderBindings(String docxPath) throws IOException {
        List<Map<String, Object>> bindings = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(docxPath); XWPFDocument doc = new XWPFDocument(fis)) {
            java.util.regex.Pattern pat = java.util.regex.Pattern.compile("Sheet「(.+?)」");
            for (XWPFParagraph p : doc.getParagraphs()) {
                String t = p.getText();
                if (t == null || !t.contains(PLACEHOLDER_PREFIX)) continue;
                var m = pat.matcher(t);
                if (!m.find()) continue;
                Map<String, Object> b = new LinkedHashMap<>();
                b.put("sheet_name", m.group(1));
                b.put("pos_index", bindings.size());
                b.put("anchor_mode", "after_para");
                b.put("pos_label", "占位-" + m.group(1));
                bindings.add(b);
            }
        }
        return bindings;
    }

    /** 向模板锚点段落后插入蓝色加粗占位段（绑定预览，annotate 流程用） */
    public static void annotateBindings(String tplPath, List<Map<String, Object>> bindings, String outPath,
                                        Integer sheetCount) throws IOException {
        XWPFDocument doc;
        try (FileInputStream fis = new FileInputStream(tplPath)) {
            doc = new XWPFDocument(fis);
        }
        int need = 0;
        for (Map<String, Object> b : bindings) {
            Object pi = b.get("pos_index") != null ? b.get("pos_index") : b.getOrDefault("anchor_idx", 0);
            need = Math.max(need, ((Number) pi).intValue() + 1);
        }
        if (sheetCount != null) need = Math.max(need, sheetCount);
        while (doc.getParagraphs().size() < need + 1) {
            XWPFParagraph p = doc.createParagraph();
            p.createRun().setText("（此处将动态插入 Sheet 表格）");
        }
        int offset = 0;
        List<Map<String, Object>> ordered = new ArrayList<>(bindings);
        ordered.sort(java.util.Comparator.comparingInt(b -> ((Number) b.getOrDefault("pos_index", b.getOrDefault("anchor_idx", 0))).intValue()));
        for (Map<String, Object> b : ordered) {
            Object pi = b.get("pos_index") != null ? b.get("pos_index") : b.getOrDefault("anchor_idx", 0);
            int posIndex = ((Number) pi).intValue();
            int paras = doc.getParagraphs().size();
            int idx = posIndex + offset;
            if (idx < 0 || idx >= paras) continue;
            String sheetName = b.get("sheet_name") != null ? String.valueOf(b.get("sheet_name")) : "";
            // 在锚点段落之后插入新段落（cursor 移过 anchor 元素后 insertNewParagraph）
            XWPFParagraph np;
            try (var cursor = doc.getParagraphs().get(idx).getCTP().newCursor()) {
                cursor.toNextToken();
                np = doc.insertNewParagraph(cursor);
            }
            XWPFRun run = np.createRun();
            run.setText(placeholderText(sheetName));
            run.setBold(true);
            run.setColor("534AB7");
            run.setFontSize(10);
            offset++;
        }
        try (FileOutputStream fos = new FileOutputStream(outPath)) {
            doc.write(fos);
        }
        doc.close();
    }

    /** 注入 Sheet 表格（样式配置全套），anchorParaIdx 非空时表格插到该段落后 */
    public static XWPFTable injectSheetTable(XWPFDocument doc, SheetData sheet, Integer anchorParaIdx,
                                             List<Double> explicitRatio, Map<String, Object> style) {
        Map<String, Object> st = resolveStyle(style);
        List<String> header = sheet.header;
        List<List<String>> rows = sheet.rows;
        int ncols = header.size();
        if (ncols == 0) return null;  // 分组 Sheet 去掉前两列后无数据列：无可注入内容

        XWPFTable table = doc.createTable(1, ncols);
        applyGridBorders(table);

        // 列宽：显式比例 > 三档模式（equal/auto/ratio）> 旧全局串兼容 > 均分
        List<Double> ratio = explicitRatio;
        String mode = String.valueOf(st.get("col_widths_mode"));
        if (ratio == null && "auto".equals(mode)) ratio = colWeightsAuto(sheet);
        if (ratio == null && "ratio".equals(mode)) {
            @SuppressWarnings("unchecked")
            Map<String, Object> cw = (Map<String, Object>) st.get("col_widths");
            ratio = parseColRatio(String.valueOf(cw.getOrDefault(sheet.name, "")), ncols);
        }
        if (ratio == null) ratio = parseColRatio(String.valueOf(st.get("col_widths_ratio")), ncols);
        int total = textWidthTwips(doc);
        int[] twips;
        double sum = ratio == null ? 0 : ratio.stream().mapToDouble(Double::doubleValue).sum();
        if (ratio != null && ratio.size() == ncols && sum > 0) {
            twips = new int[ncols];
            for (int j = 0; j < ncols; j++) twips[j] = (int) (total * ratio.get(j) / sum);
        } else {
            twips = new int[ncols];
            java.util.Arrays.fill(twips, total / ncols);
        }

        String vAlign = String.valueOf(st.get("v_align"));
        String hAlign = String.valueOf(st.get("h_align"));
        double fontSize = ((Number) st.get("font_size")).doubleValue();
        String fontName = String.valueOf(st.get("font_name"));

        for (int j = 0; j < ncols; j++) {
            setCellText(table.getRow(0).getCell(j), header.get(j), true, fontSize, hAlign, fontName, vAlign);
        }
        for (List<String> rowData : rows) {
            XWPFTableRow row = table.createRow();
            for (int j = 0; j < ncols; j++) {
                setCellText(row.getCell(j), j < rowData.size() ? rowData.get(j) : "", false, fontSize, hAlign, fontName, vAlign);
            }
        }
        if ("atLeast".equals(st.get("row_height_mode"))) {
            double ptv = ((Number) st.get("row_height_pt")).doubleValue();
            for (XWPFTableRow row : table.getRows()) setRowMinHeight(row, ptv);
        }
        for (XWPFTableRow row : table.getRows()) {
            for (int j = 0; j < ncols; j++) {
                var tcPr = row.getCell(j).getCTTc().isSetTcPr()
                        ? row.getCell(j).getCTTc().getTcPr() : row.getCell(j).getCTTc().addNewTcPr();
                var tcW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
                tcW.setW(BigInteger.valueOf(twips[j]));
                tcW.setType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth.DXA);
            }
        }

        if (anchorParaIdx != null) moveTableAfterParagraph(doc, table, anchorParaIdx);
        return table;
    }

    /** 渲染统计（对齐 Python stats 键） */
    public static class RenderStats {
        public int templateParaCount, templateTableCount, outputTableCount;
        public double textWidthInch;
        public List<Map<String, Object>> injected = new ArrayList<>();
        public List<String> skippedEmpty = new ArrayList<>();
        public List<String> notFound = new ArrayList<>();
    }

    /**
     * 渲染主流程（对齐 Python render_template）：占位段优先注入表格并删除占位段；
     * 无数据/缺失 Sheet 删除占位段；保存前宽松终扫兜底删除全部残留占位段。
     * groupKey 非空时启用批量模式（虚拟 sheet，前两列分组列去掉）。
     */
    public static RenderStats renderTemplate(String tplPath, String xlsxPath, List<Map<String, Object>> bindings,
                                             String outPath, String[] groupKey, Map<String, Object> style) throws IOException {
        XWPFDocument doc;
        try (FileInputStream fis = new FileInputStream(tplPath)) {
            doc = new XWPFDocument(fis);
        }
        int textWidthTw = textWidthTwips(doc);

        Map<String, SheetData> sheetMap = new LinkedHashMap<>();
        List<SheetData> sheetsInOrder = new ArrayList<>();
        if (groupKey != null) {
            GroupedResult grouped = readExcelGrouped(xlsxPath);
            String k = groupKey[0] + "\n" + groupKey[1];
            for (String name : grouped.sheets) {
                SheetData vs = grouped.virtualSheets.get(name).get(k);
                if (vs != null) { sheetMap.put(name, vs); sheetsInOrder.add(vs); }
            }
        } else {
            for (SheetData s : readExcelSheets(xlsxPath)) { sheetMap.put(s.name, s); sheetsInOrder.add(s); }
        }

        RenderStats stats = new RenderStats();
        stats.templateParaCount = doc.getParagraphs().size();
        stats.templateTableCount = doc.getTables().size();
        stats.textWidthInch = Math.round(textWidthTw / 1440.0 * 100.0) / 100.0;  // twips -> inch

        Map<String, Integer> occurrence = new LinkedHashMap<>();  // 同一 Sheet 多处插入的出现计数
        // 占位段延迟统一删除：表格是独立 body 元素，注入不影响段落索引（稳定）；
        // 全部注入完成后合并终扫索引倒序一次删除（removeParagraphAt 走 removeBodyElement 同步内部缓存）
        List<Integer> placeholdersToRemove = new ArrayList<>();

        for (Map<String, Object> binding : bindings) {
            String sheetName = String.valueOf(binding.get("sheet_name"));
            SheetData sheet = sheetMap.get(sheetName);
            int occ = occurrence.getOrDefault(sheetName, 0);

            int phIdx = findPlaceholderPara(doc, sheetName, occ);
            if (sheet == null) {
                stats.notFound.add(sheetName);
                if (phIdx >= 0) { placeholdersToRemove.add(phIdx); occurrence.put(sheetName, occ + 1); }
                continue;
            }
            if (sheet.rows.isEmpty()) {
                stats.skippedEmpty.add(sheetName);
                if (phIdx >= 0) { placeholdersToRemove.add(phIdx); occurrence.put(sheetName, occ + 1); }
                continue;
            }

            if (phIdx >= 0) {
                // 占位段优先：表格插在占位段之后，占位段延迟统一删除
                injectSheetTable(doc, sheet, phIdx, null, style);
                placeholdersToRemove.add(phIdx);
                occurrence.put(sheetName, occ + 1);
            } else {
                // 无占位段 fallback：文档末尾
                injectSheetTable(doc, sheet, null, null, style);
            }
            Map<String, Object> inj = new LinkedHashMap<>();
            inj.put("sheet", sheetName);
            inj.put("cols", sheet.header.size());
            inj.put("rows", sheet.rows.size() + 1);
            stats.injected.add(inj);
        }

        // 统一删除占位段 + 终扫兜底（宽松匹配「【Sheet」+「将在此处展示」，倒序删除防索引位移）
        List<XWPFParagraph> paras = doc.getParagraphs();
        for (int i = 0; i < paras.size(); i++) {
            String t = paras.get(i).getText();
            if (t != null && t.contains(PLACEHOLDER_PREFIX) && t.contains("将在此处展示")
                    && !placeholdersToRemove.contains(i)) {
                placeholdersToRemove.add(i);
            }
        }
        placeholdersToRemove.sort(java.util.Comparator.reverseOrder());
        for (int i : placeholdersToRemove) removeParagraphAt(doc, i);
        if (!placeholdersToRemove.isEmpty()) {
            System.out.println("[RENDER] 已删除占位段 " + placeholdersToRemove.size() + " 处");
        }

        try (FileOutputStream fos = new FileOutputStream(outPath)) {
            doc.write(fos);
        }
        doc.close();
        try (FileInputStream fis = new FileInputStream(outPath); XWPFDocument saved = new XWPFDocument(fis)) {
            stats.outputTableCount = saved.getTables().size();
        }
        return stats;
    }
}