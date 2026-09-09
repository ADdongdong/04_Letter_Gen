package com.xidian.lettergen;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.util.*;
import java.util.stream.Collectors;

import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * 模板配置存储（templates_store：index.json + <id>/word.docx + <id>/excel.xlsx）。
 * 存储结构与 Python 版完全兼容（同一目录结构，Python/Java 后端可互换使用同一数据）。
 */
public class TemplateStore {

    public final File baseDir;        // backend/java（与 Python 版的 backend/python 对应，存储互相独立）
    public final File templatesDir;   // templates_store/
    public final File templatesIndex; // templates_store/index.json
    public final File uploadDir;      // uploads/
    public final File outputDir;      // outputs/
    public final File currentTemplate; // uploads/templates/current.docx
    public final File defaultTpl;     // backend/static-docs/demo_template.docx
    public final File lettersRegistry; // data/letters_registry.json
    public final File staticDocsDir;  // backend/static-docs

    @SuppressWarnings("unchecked")
    public TemplateStore() {
        try {
            baseDir = new File(".").getCanonicalFile();  // 运行工作目录 = backend/java
        } catch (IOException e) {
            throw new IllegalStateException("无法解析工作目录", e);
        }
        templatesDir = new File(baseDir, "templates_store");
        templatesIndex = new File(templatesDir, "index.json");
        uploadDir = new File(baseDir, "uploads");
        outputDir = new File(baseDir, "outputs");
        currentTemplate = new File(uploadDir, "templates/current.docx");
        defaultTpl = new File(baseDir.getParentFile(), "static-docs/demo_template.docx");
        lettersRegistry = new File(baseDir.getParentFile(), "python/data/letters_registry.json");
        staticDocsDir = new File(baseDir.getParentFile(), "static-docs");
        try {
            Files.createDirectories(templatesDir.toPath());
            Files.createDirectories(uploadDir.toPath());
            Files.createDirectories(outputDir.toPath());
            Files.createDirectories(new File(uploadDir, "templates").toPath());
            Files.createDirectories(staticDocsDir.toPath());
        } catch (IOException e) {
            throw new IllegalStateException("无法创建存储目录", e);
        }
    }

    /** 当前生效模板：用户上传的优先，否则用内置默认模板 */
    public File currentTemplate() {
        if (currentTemplate.exists() && currentTemplate.length() > 0) return currentTemplate;
        return defaultTpl;
    }

    private final ObjectMapper mapper = new ObjectMapper();

    /** index.json 读取（List<Map>，Jackson 结构与 Python json 对齐） */
    @SuppressWarnings("unchecked")
    public synchronized List<Map<String, Object>> loadIndex() {
        if (!templatesIndex.exists()) return new ArrayList<>();
        try {
            return mapper.readValue(templatesIndex, List.class);
        } catch (Exception e) {
            return new ArrayList<>();
        }
    }

    public synchronized void saveIndex(List<Map<String, Object>> items) throws IOException {
        Files.writeString(templatesIndex.toPath(),
                mapper.writerWithDefaultPrettyPrinter().writeValueAsString(items), StandardCharsets.UTF_8);
    }

    public Map<String, Object> findTemplate(List<Map<String, Object>> items, String id) {
        return items.stream().filter(x -> id.equals(x.get("id"))).findFirst().orElse(null);
    }

    /** 函证登记清单：[{audit, confirm, letter_no}] */
    @SuppressWarnings("unchecked")
    public List<Map<String, Object>> loadLettersRegistry() {
        if (!lettersRegistry.exists()) return new ArrayList<>();
        try {
            return mapper.readValue(lettersRegistry, List.class);
        } catch (Exception e) {
            return new ArrayList<>();
        }
    }

    /** 组合与函证登记数据精确匹配：[{audit, confirm, letter_no, matched}] */
    public List<Map<String, Object>> matchLetters(List<String[]> groups, List<Map<String, Object>> letters) {
        Map<String, String> index = new HashMap<>();
        for (Map<String, Object> it : letters) {
            String a = String.valueOf(it.getOrDefault("audit", "")).trim();
            String c = String.valueOf(it.getOrDefault("confirm", "")).trim();
            String no = it.get("letter_no") == null ? "" : String.valueOf(it.get("letter_no")).trim();
            index.put(a + "\n" + c, no);
        }
        List<Map<String, Object>> result = new ArrayList<>();
        for (String[] g : groups) {
            String no = index.getOrDefault(g[0] + "\n" + g[1], "");
            Map<String, Object> m = new LinkedHashMap<>();
            m.put("audit", g[0]); m.put("confirm", g[1]);
            m.put("letter_no", no.isEmpty() ? null : no);
            m.put("matched", !no.isEmpty());
            result.add(m);
        }
        return result;
    }

    /** 清洗 Windows 文件名非法字符 /\\:*?"<>| → _（对齐 Python _safe_filename） */
    public static String safeFilename(String name) {
        String cleaned = name == null ? "" : name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        return cleaned.isEmpty() ? "未命名" : cleaned;
    }

    /** Python str() 风格的空单元格消息定位：Sheet名 第N行（列名）；超 20 处截断 */
    public static String formatEmptyCells(List<Map<String, Object>> emptyCells) {
        String joined = emptyCells.stream().limit(20)
                .map(c -> "%s 第%d行（%s）".formatted(c.get("sheet"), ((Number) c.get("row")).intValue(), c.get("col")))
                .collect(Collectors.joining("；"));
        if (emptyCells.size() > 20) joined += "；等共 " + emptyCells.size() + " 处";
        return joined;
    }
}
