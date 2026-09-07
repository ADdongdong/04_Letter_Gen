package com.xidian.lettergen;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

/**
 * 制函服务全部 REST 端点（功能对齐 Python 版 backend/python/app.py，前端共用零改动）。
 */
@RestController
public class ApiController {

    private final TemplateStore store = new TemplateStore();
    private final ObjectMapper mapper = new ObjectMapper();
    // 显式禁用系统代理 + 固定 HTTP/1.1：
    // - NO_PROXY 防止 Clash 等系统代理劫持内部请求
    // - Java HttpClient 默认 HTTP/2 会向明文服务发 h2c 升级请求，OO 的 nginx 返回 HTML 错误页
    private final HttpClient http = HttpClient.newBuilder()
            .proxy(HttpClient.Builder.NO_PROXY)
            .version(HttpClient.Version.HTTP_1_1)
            .build();

    // OO forcesave 回调记录（全局，供前端轮询判断保存是否完成）与活跃编辑器 key
    private volatile String lastSaveKey = null;
    private volatile double lastSaveTs = 0.0;
    private volatile String activeDocKey = null;
    // OO Document Server 宿主机访问地址（docker 端口映射 8080->80）
    // 必须用 127.0.0.1 而非 localhost：Java HttpClient 解析 localhost 时 IPv6 优先，
    // 本机 [::1]:8080 被 wslrelay.exe（WSL 端口中继）占用且返回 HTML，会导致 forcesave 收到非 JSON
    private static final String OO_PUBLIC_URL = "http://127.0.0.1:8080";

    private static final String GROUP_AUDIT = "被审计单位";
    private static final String GROUP_CONFIRM = "被询证单位";

    private ResponseEntity<Map<String, Object>> err(String msg, HttpStatus status) {
        return ResponseEntity.status(status).body(Map.of("error", msg));
    }

    // ============ 根路径 ============
    @GetMapping("/")
    public Map<String, Object> index() {
        return Map.of("service", "lettergen-java", "version", "1.0.0");
    }

    // ============ 上传模板 ============
    @PostMapping("/api/upload_template")
    public ResponseEntity<?> uploadTemplate(@RequestParam("file") MultipartFile file) throws IOException {
        String name = StringUtils.hasText(file.getOriginalFilename()) ? file.getOriginalFilename() : "";
        if (!name.toLowerCase().endsWith(".docx")) {
            return err("请上传 .docx 模板", HttpStatus.BAD_REQUEST);
        }
        String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        File path = new File(store.uploadDir, "tpl_" + uid + ".docx");
        file.transferTo(path);

        // 同步为当前模板：此后 /api/oo/getTemplate 与渲染均使用它
        Files.copy(path.toPath(), store.currentTemplate.toPath(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        System.out.println("[TPL-UP] 模板已上传并生效: " + name + " -> " + store.currentTemplate);

        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("path", path.getAbsolutePath());
        resp.put("uid", uid);
        resp.put("name", name);
        resp.put("size_kb", Math.round(path.length() / 102.4) / 10.0);
        try (FileInputStream fis = new FileInputStream(path); XWPFDocument doc = new XWPFDocument(fis)) {
            List<Map<String, Object>> paras = new ArrayList<>();
            int pi = 0;
            for (var p : doc.getParagraphs()) {
                String t = p.getText() == null ? "" : p.getText();
                if (!t.isBlank()) {
                    Map<String, Object> pm = new LinkedHashMap<>();
                    pm.put("idx", pi); pm.put("text", t.substring(0, Math.min(30, t.length())));
                    paras.add(pm);
                }
                pi++;
            }
            List<Map<String, Object>> tables = new ArrayList<>();
            int ti = 0;
            for (XWPFTable t : doc.getTables()) {
                Map<String, Object> tm = new LinkedHashMap<>();
                tm.put("idx", ti);
                tm.put("rows", t.getRows().size());
                tm.put("cols", t.getRows().isEmpty() ? 0 : t.getRow(0).getTableCells().size());
                tm.put("first_cell", t.getRows().isEmpty() || t.getRow(0).getTableCells().isEmpty()
                        ? "" : t.getRow(0).getCell(0).getText().substring(0, Math.min(20, t.getRow(0).getCell(0).getText().length())));
                tables.add(tm);
                ti++;
            }
            double tw = RenderEngine.textWidthTwips(doc) / 1440.0;
            resp.put("para_count", doc.getParagraphs().size());
            resp.put("table_count", doc.getTables().size());
            resp.put("text_width_inch", Math.round(tw * 100.0) / 100.0);
            resp.put("paras", paras);
            resp.put("tables", tables);
        }
        return ResponseEntity.ok(resp);
    }

    // ============ 函证登记清单 ============
    @GetMapping("/api/letters")
    public Map<String, Object> lettersList() {
        return Map.of("letters", store.loadLettersRegistry());
    }

    // ============ 上传 Excel（校验 + 分组预览 + 函证编号匹配） ============
    @PostMapping("/api/upload_excel")
    public ResponseEntity<?> uploadExcel(@RequestParam("file") MultipartFile file) throws IOException {
        String name = StringUtils.hasText(file.getOriginalFilename()) ? file.getOriginalFilename() : "";
        if (!(name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xlsm"))) {
            return err("请上传 .xlsx 文件", HttpStatus.BAD_REQUEST);
        }
        String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        File path = new File(store.uploadDir, "excel_" + uid + ".xlsx");
        file.transferTo(path);

        List<RenderEngine.SheetData> sheets = RenderEngine.readExcelSheets(path.getAbsolutePath());
        RenderEngine.GroupedResult grouped = RenderEngine.readExcelGrouped(path.getAbsolutePath());
        if (!grouped.emptyCells.isEmpty()) {
            return err("检测到 %d 处「被审计单位/被询证单位」为空，请补充后重新上传：%s"
                    .formatted(grouped.emptyCells.size(), TemplateStore.formatEmptyCells(grouped.emptyCells)),
                    HttpStatus.BAD_REQUEST);
        }

        List<Map<String, Object>> sheetInfos = new ArrayList<>();
        for (RenderEngine.SheetData s : sheets) {
            Map<String, Object> si = new LinkedHashMap<>();
            si.put("name", s.name);
            si.put("header", s.header);
            si.put("row_count", s.rows.size());
            si.put("preview", s.rows.subList(0, Math.min(5, s.rows.size())));
            si.put("is_grouped", s.header.size() >= 2
                    && GROUP_AUDIT.equals(s.header.get(0).trim()) && GROUP_CONFIRM.equals(s.header.get(1).trim()));
            sheetInfos.add(si);
        }
        List<Map<String, Object>> groups = new ArrayList<>();
        for (String[] g : grouped.groups) {
            groups.add(Map.of("audit", g[0], "confirm", g[1]));
        }
        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("path", path.getAbsolutePath());
        resp.put("uid", uid);
        resp.put("sheets", sheetInfos);
        resp.put("batch_mode", !grouped.groups.isEmpty());
        resp.put("groups", groups);
        resp.put("match_results", store.matchLetters(grouped.groups, store.loadLettersRegistry()));
        return ResponseEntity.ok(resp);
    }

    // ============ 模板锚点（旧流程动态锚点，保留兼容） ============
    @GetMapping("/api/template_anchors")
    public ResponseEntity<?> templateAnchors() {
        File tpl = store.currentTemplate();
        if (!tpl.exists()) return err("默认模板不存在: " + tpl, HttpStatus.NOT_FOUND);
        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("template", tpl.getAbsolutePath());
        resp.put("anchors", List.of());
        return ResponseEntity.ok(resp);
    }

    // ============ 绑定回写（模板上显示绑定关系，生成蓝色占位段预览 docx） ============
    @PostMapping("/api/bind_annotate")
    @SuppressWarnings("unchecked")
    public ResponseEntity<?> bindAnnotate(@RequestBody Map<String, Object> data) throws IOException {
        List<Map<String, Object>> bindings = (List<Map<String, Object>>) data.getOrDefault("bindings", List.of());
        String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        File outPath = new File(store.outputDir, "annotated_" + uid + ".docx");
        Integer sheetCount = data.get("sheet_count") == null ? null : ((Number) data.get("sheet_count")).intValue();
        RenderEngine.annotateBindings(getCurrentTemplate().getAbsolutePath(), bindings,
                outPath.getAbsolutePath(), sheetCount);
        // 复制到静态目录，便于 OnlyOffice 通过 http 加载
        File staticDir = store.staticDocsDir;
        Files.copy(outPath.toPath(), new File(staticDir, outPath.getName()).toPath(),
                java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("path", outPath.getAbsolutePath());
        resp.put("download_url", "/api/download/" + outPath.getName());
        resp.put("bindings", bindings);
        return ResponseEntity.ok(resp);
    }

    private File getCurrentTemplate() {
        return store.currentTemplate();
    }

    // ============ 单发渲染（旧 demo 流程，保留兼容） ============
    @PostMapping("/api/render")
    @SuppressWarnings("unchecked")
    public ResponseEntity<?> render(@RequestBody Map<String, Object> data) {
        String tplPath = data.get("tpl_path") != null ? String.valueOf(data.get("tpl_path"))
                : getCurrentTemplate().getAbsolutePath();
        String excelPath = (String) data.get("excel_path");
        List<Map<String, Object>> bindings = (List<Map<String, Object>>) data.getOrDefault("bindings", List.of());
        if (!new File(tplPath).exists()) return err("模板不存在: " + tplPath, HttpStatus.BAD_REQUEST);
        if (excelPath == null || !new File(excelPath).exists()) return err("请先上传 Excel", HttpStatus.BAD_REQUEST);
        if (bindings.isEmpty()) return err("请至少添加一个绑定", HttpStatus.BAD_REQUEST);
        try {
            String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
            File out = new File(store.outputDir, "result_" + uid + ".docx");
            var stats = RenderEngine.renderTemplate(tplPath, excelPath, bindings, out.getAbsolutePath(), null, null);
            Map<String, Object> resp = new LinkedHashMap<>(new ObjectMapper().convertValue(stats, Map.class));
            resp.put("download_url", "/api/download/" + out.getName());
            return ResponseEntity.ok(resp);
        } catch (Exception e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // ============ 旧批量渲染接口（保留兼容；制函页走 /api/generate） ============
    @PostMapping("/api/render_batch")
    @SuppressWarnings("unchecked")
    public ResponseEntity<?> renderBatch(@RequestBody Map<String, Object> data) {
        String tplPath = data.get("tpl_path") != null ? String.valueOf(data.get("tpl_path"))
                : getCurrentTemplate().getAbsolutePath();
        String excelPath = (String) data.get("excel_path");
        List<Map<String, Object>> bindings = (List<Map<String, Object>>) data.getOrDefault("bindings", List.of());
        if (!new File(tplPath).exists()) return err("模板不存在: " + tplPath, HttpStatus.BAD_REQUEST);
        if (excelPath == null || !new File(excelPath).exists()) return err("请先上传 Excel", HttpStatus.BAD_REQUEST);
        if (bindings.isEmpty()) return err("请至少添加一个绑定", HttpStatus.BAD_REQUEST);
        try {
            RenderEngine.GroupedResult grouped = RenderEngine.readExcelGrouped(excelPath);
            if (grouped.groups.isEmpty()) {
                return err("未检测到批量数据：Excel 中没有表头前两列为「被审计单位/被询证单位」的 Sheet", HttpStatus.BAD_REQUEST);
            }
            if (!grouped.emptyCells.isEmpty()) {
                return err("检测到 %d 处「被审计单位/被询证单位」为空，请补充后重新上传：%s"
                        .formatted(grouped.emptyCells.size(), TemplateStore.formatEmptyCells(grouped.emptyCells)),
                        HttpStatus.BAD_REQUEST);
            }
            return doBatchRender(tplPath, excelPath, bindings, grouped, null, "batch");
        } catch (Exception e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    /** 批量渲染共用：N 封 docx 打包 ZIP。prefix：batch|gen */
    private ResponseEntity<?> doBatchRender(String tplPath, String excelPath, List<Map<String, Object>> bindings,
                                            RenderEngine.GroupedResult grouped, Map<String, Object> style,
                                            String prefix) throws IOException {
        String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        File batchDir = new File(store.outputDir, prefix + "_" + uid);
        batchDir.mkdirs();
        List<Map<String, Object>> files = new ArrayList<>();
        for (String[] g : grouped.groups) {
            String safe = TemplateStore.safeFilename(g[0] + "-" + g[1]);
            File out = new File(batchDir, safe + ".docx");
            var stats = RenderEngine.renderTemplate(tplPath, excelPath, bindings, out.getAbsolutePath(),
                    new String[]{g[0], g[1]}, style);
            Map<String, Object> f = new LinkedHashMap<>();
            f.put("file", safe + ".docx");
            f.put("audit", g[0]); f.put("confirm", g[1]);
            f.put("tables", stats.outputTableCount);
            f.put("skipped_empty", stats.skippedEmpty);
            files.add(f);
            System.out.println("[BATCH] " + g[0] + " - " + g[1] + ": tables=" + stats.outputTableCount);
        }
        File zip = new File(store.outputDir, prefix + "_" + uid + ".zip");
        try (ZipOutputStream zf = new ZipOutputStream(new FileOutputStream(zip))) {
            for (Map<String, Object> f : files) {
                zf.putNextEntry(new ZipEntry((String) f.get("file")));
                Files.copy(new File(batchDir, (String) f.get("file")).toPath(), zf);
                zf.closeEntry();
            }
        }
        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("count", files.size());
        resp.put("files", files);
        resp.put("download_url", "/api/download/" + zip.getName());
        return ResponseEntity.ok(resp);
    }

    // ============ 下载 ============
    @GetMapping("/api/download/{fname}")
    public ResponseEntity<?> download(@PathVariable String fname) {
        File path = new File(store.outputDir, fname);
        if (!path.exists()) return err("文件不存在", HttpStatus.NOT_FOUND);
        boolean isZip = fname.toLowerCase().endsWith(".zip");
        return fileResponse(path, isZip ? "渲染结果.zip" : "渲染结果.docx", true);
    }

    private ResponseEntity<?> fileResponse(File path, String downloadName, boolean asAttachment) {
        try {
            byte[] body = Files.readAllBytes(path.toPath());
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.parseMediaType(
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"));
            headers.set(HttpHeaders.CACHE_CONTROL, "no-store, must-revalidate");
            if (asAttachment) {
                headers.setContentDispositionFormData("attachment",
                        java.net.URLEncoder.encode(downloadName, StandardCharsets.UTF_8));
            }
            return new ResponseEntity<>(body, headers, HttpStatus.OK);
        } catch (IOException e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // ============ 模板配置存储 ============
    @GetMapping("/api/templates")
    public Map<String, Object> templatesList() {
        return Map.of("templates", store.loadIndex());
    }

    @PostMapping("/api/templates/save")
    public ResponseEntity<?> templatesSave(@RequestBody Map<String, Object> data) {
        String name = data.get("name") == null ? "" : String.valueOf(data.get("name")).trim();
        String excelPath = data.get("excel_path") == null ? null : String.valueOf(data.get("excel_path"));
        String tid = data.get("id") == null ? null : String.valueOf(data.get("id"));
        if (name.isEmpty()) return err("请填写模板名称", HttpStatus.BAD_REQUEST);
        if (excelPath != null && !new File(excelPath).exists()) {
            return err("上传的 Excel 不存在", HttpStatus.BAD_REQUEST);
        }
        if (tid == null && excelPath == null) {
            return err("请先上传 Excel（用于定义 Sheet 结构，可以没有数据）", HttpStatus.BAD_REQUEST);
        }
        List<Map<String, Object>> items = store.loadIndex();
        for (Map<String, Object> it : items) {
            if (name.equals(it.get("name")) && !tidIsSame(it, tid)) {
                return err("模板名称「" + name + "」已存在", HttpStatus.BAD_REQUEST);
            }
        }
        String now = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        Map<String, Object> item;
        if (tid != null) {
            item = store.findTemplate(items, tid);
            if (item == null) return err("模板不存在", HttpStatus.NOT_FOUND);
            item.put("name", name);
            item.put("updated_at", now);
        } else {
            tid = "tpl_" + UUID.randomUUID().toString().replace("-", "").substring(0, 8);
            item = new LinkedHashMap<>();
            item.put("id", tid); item.put("name", name);
            item.put("created_at", now); item.put("updated_at", now);
            item.put("sheets", new ArrayList<>());
            items.add(item);
        }
        File tplDir = new File(store.templatesDir, tid);
        tplDir.mkdirs();
        try {
            Files.copy(getCurrentTemplate().toPath(), new File(tplDir, "word.docx").toPath(),
                    java.nio.file.StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
        // Excel：编辑模式未重新上传时沿用该模板已存的 excel.xlsx 与 sheets（不覆盖）
        if (excelPath != null) {
            File dst = new File(tplDir, "excel.xlsx");
            try {
                if (!(dst.exists() && Files.isSameFile(new File(excelPath).toPath(), dst.toPath()))) {
                    Files.copy(new File(excelPath).toPath(), dst.toPath(),
                            java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                }
                List<String> names = new ArrayList<>();
                for (RenderEngine.SheetData s : RenderEngine.readExcelSheets(excelPath)) names.add(s.name);
                item.put("sheets", names);
            } catch (IOException e) {
                item.put("sheets", new ArrayList<>());
            }
        }
        // 表格样式：仅请求携带 style 字段时更新，经 resolveStyle 校验合并
        if (data.containsKey("style")) {
            @SuppressWarnings("unchecked")
            Map<String, Object> raw = (Map<String, Object>) data.get("style");
            item.put("style", RenderEngine.resolveStyle(raw));
        }
        try {
            store.saveIndex(items);
        } catch (IOException e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
        System.out.println("[TPL-SAVE] 模板配置已保存: id=" + tid + " name=" + name + " sheets=" + item.get("sheets"));
        Map<String, Object> resp = new LinkedHashMap<>();
        resp.put("id", tid); resp.put("name", name);
        resp.put("sheets", item.get("sheets")); resp.put("style", item.get("style"));
        return ResponseEntity.ok(resp);
    }

    private boolean tidIsSame(Map<String, Object> it, String tid) {
        return tid != null && tid.equals(it.get("id"));
    }

    @DeleteMapping("/api/templates/{tid}")
    public ResponseEntity<?> templatesDelete(@PathVariable String tid) throws IOException {
        List<Map<String, Object>> items = store.loadIndex();
        Map<String, Object> item = store.findTemplate(items, tid);
        if (item == null) return err("模板不存在", HttpStatus.NOT_FOUND);
        items.remove(item);
        store.saveIndex(items);
        File tplDir = new File(store.templatesDir, tid);
        if (tplDir.isDirectory()) deleteRecursively(tplDir);
        System.out.println("[TPL-DEL] 模板配置已删除: id=" + tid + " name=" + item.get("name"));
        return ResponseEntity.ok(Map.of("ok", true));
    }

    private static void deleteRecursively(File f) {
        File[] children = f.listFiles();
        if (children != null) for (File c : children) deleteRecursively(c);
        f.delete();
    }

    /** 模板 word 下载（配置页修改模式时 OnlyOffice 经 host.docker.internal 加载用） */
    @GetMapping("/api/templates/{tid}/word")
    public ResponseEntity<?> templatesWord(@PathVariable String tid) {
        File path = new File(store.templatesDir, tid + "/word.docx");
        if (!path.exists()) return err("模板不存在", HttpStatus.NOT_FOUND);
        return fileResponse(path, "template.docx", false);
    }

    // ============ 批量制函（制函页：下拉选模板 + 上传批量 Excel，模板即定义） ============
    @PostMapping("/api/generate")
    @SuppressWarnings("unchecked")
    public ResponseEntity<?> generate(@RequestBody(required = false) Map<String, Object> data) {
        if (data == null) data = Map.of();
        String templateId = data.get("template_id") == null ? null : String.valueOf(data.get("template_id"));
        String excelPath = data.get("excel_path") == null ? null : String.valueOf(data.get("excel_path"));
        if (templateId == null) return err("请选择制函模板", HttpStatus.BAD_REQUEST);
        if (excelPath == null || !new File(excelPath).exists()) return err("请先上传 Excel", HttpStatus.BAD_REQUEST);

        Map<String, Object> item = store.findTemplate(store.loadIndex(), templateId);
        if (item == null) return err("模板不存在", HttpStatus.NOT_FOUND);
        File tplWord = new File(store.templatesDir, templateId + "/word.docx");
        if (!tplWord.exists()) return err("模板 Word 文件缺失", HttpStatus.NOT_FOUND);

        RenderEngine.GroupedResult grouped;
        try {
            grouped = RenderEngine.readExcelGrouped(excelPath);
        } catch (Exception e) {
            return err("解析 Excel 失败: " + e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
        if (grouped.groups.isEmpty()) {
            return err("请上传批量格式 Excel：各 Sheet 表头前两列需为「被审计单位」「被询证单位」", HttpStatus.BAD_REQUEST);
        }
        if (!grouped.emptyCells.isEmpty()) {
            return err("检测到 %d 处「被审计单位/被询证单位」为空，请补充后重新上传：%s"
                    .formatted(grouped.emptyCells.size(), TemplateStore.formatEmptyCells(grouped.emptyCells)),
                    HttpStatus.BAD_REQUEST);
        }
        var matches = store.matchLetters(grouped.groups, store.loadLettersRegistry());
        List<String[]> matched = new ArrayList<>();
        List<Map<String, Object>> unmatched = new ArrayList<>();
        for (Map<String, Object> m : matches) {
            if (Boolean.TRUE.equals(m.get("matched"))) {
                matched.add(new String[]{(String) m.get("audit"), (String) m.get("confirm"), (String) m.get("letter_no")});
            } else {
                unmatched.add(Map.of("audit", m.get("audit"), "confirm", m.get("confirm")));
            }
        }
        if (matched.isEmpty()) {
            return err("Excel 中所有函证在系统中均未登记（缺少对应的函证编号），无法制函。请先在函证系统完成登记。",
                    HttpStatus.BAD_REQUEST);
        }
        List<Map<String, Object>> bindings;
        try {
            bindings = RenderEngine.extractPlaceholderBindings(tplWord.getAbsolutePath());
        } catch (IOException e) {
            return err(String.valueOf(e.getMessage()), HttpStatus.INTERNAL_SERVER_ERROR);
        }
        if (bindings.isEmpty()) {
            return err("模板中未找到任何占位段（【Sheet「xxx」...】），请先在配置页完成标注", HttpStatus.BAD_REQUEST);
        }
        List<String> sheetMissing = bindings.stream()
                .map(b -> String.valueOf(b.get("sheet_name")))
                .filter(n -> !grouped.sheets.contains(n))
                .distinct().sorted().toList();

        @SuppressWarnings("unchecked")
        Map<String, Object> style = (Map<String, Object>) item.get("style");
        String uid = UUID.randomUUID().toString().replace("-", "").substring(0, 8);
        File batchDir = new File(store.outputDir, "gen_" + uid);
        batchDir.mkdirs();
        List<Map<String, Object>> files = new ArrayList<>();
        try {
            for (String[] m : matched) {
                String safe = TemplateStore.safeFilename(m[0] + "-" + m[1]);
                File out = new File(batchDir, safe + ".docx");
                var stats = RenderEngine.renderTemplate(tplWord.getAbsolutePath(), excelPath, bindings,
                        out.getAbsolutePath(), new String[]{m[0], m[1]}, style);
                Map<String, Object> f = new LinkedHashMap<>();
                f.put("file", safe + ".docx");
                f.put("audit", m[0]); f.put("confirm", m[1]); f.put("letter_no", m[2]);
                f.put("tables", stats.outputTableCount);
                f.put("skipped_empty", stats.skippedEmpty);
                f.put("not_found", stats.notFound);
                files.add(f);
                System.out.println("[GENERATE] " + m[0] + " - " + m[1] + " (" + m[2] + "): tables=" + stats.outputTableCount);
            }
            File zip = new File(store.outputDir, "gen_" + uid + ".zip");
            try (ZipOutputStream zf = new ZipOutputStream(new FileOutputStream(zip))) {
                for (Map<String, Object> f : files) {
                    zf.putNextEntry(new ZipEntry((String) f.get("file")));
                    Files.copy(new File(batchDir, (String) f.get("file")).toPath(), zf);
                    zf.closeEntry();
                }
            }
            Map<String, Object> resp = new LinkedHashMap<>();
            resp.put("count", files.size());
            resp.put("files", files);
            resp.put("unmatched", unmatched);
            resp.put("sheet_missing", sheetMissing);
            resp.put("download_url", "/api/download/" + zip.getName());
            return ResponseEntity.ok(resp);
        } catch (Exception e) {
            e.printStackTrace();  // 定位用：完整堆栈
            return err(e.getMessage() == null ? (e.getClass().getSimpleName() + "（无消息，见后端日志堆栈）") : e.getMessage(),
                    HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    // ============ OnlyOffice 集成 ============

    /** OO 下载当前模板（新增模式 docUrl） */
    @GetMapping("/api/oo/getTemplate")
    public ResponseEntity<?> ooGetTemplate() {
        System.out.println("[OO-DL][" + LocalDateTime.now() + "] OnlyOffice 请求模板下载");
        File tpl = getCurrentTemplate();
        if (!tpl.exists()) return err("模板不存在: " + tpl, HttpStatus.NOT_FOUND);
        return fileResponse(tpl, "demo_template.docx", false);
    }

    /** OO 保存回调（status=1 登记/2,4,6,7 回写 current.docx；旧 key 的回写跳过防竞态） */
    @PostMapping("/api/oo/callback")
    public ResponseEntity<?> ooCallback(@RequestBody(required = false) Map<String, Object> data) {
        if (data == null) data = Map.of();
        Object statusObj = data.get("status");
        int status = statusObj instanceof Number n ? n.intValue() : -1;
        String key = data.get("key") == null ? "" : String.valueOf(data.get("key"));
        System.out.println("[OO-CB][" + LocalDateTime.now() + "] status=" + status + " key="
                + key.substring(0, Math.min(24, key.length())));
        if (status == 1 && !key.isEmpty()) activeDocKey = key;
        if (status == 2 || status == 4 || status == 6 || status == 7) {
            if (!key.equals(activeDocKey)) {
                System.out.println("[OO-CB] [SKIP] 跳过旧 key 回写: "
                        + key.substring(0, Math.min(24, key.length())));
                return ResponseEntity.ok(Map.of("key", key, "error", 0));
            }
            String url = data.get("url") == null ? null : String.valueOf(data.get("url"));
            if (url != null && !url.isEmpty()) {
                try {
                    // OO 给的 url 是容器内地址（127.0.0.1:8000），宿主机需替换为映射地址
                    String dlUrl = url.replace("http://127.0.0.1:8000", OO_PUBLIC_URL);
                    File tmp = File.createTempFile("oo_", ".docx", store.uploadDir);
                    http.send(HttpRequest.newBuilder(URI.create(dlUrl)).GET().build(),
                            HttpResponse.BodyHandlers.ofFile(tmp.toPath()));
                    if (tmp.length() > 0) {
                        Files.copy(tmp.toPath(), store.currentTemplate.toPath(),
                                java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                        lastSaveKey = key;
                        lastSaveTs = System.currentTimeMillis() / 1000.0;
                        System.out.println("[OO-CB] [OK] 模板已回写: " + store.currentTemplate
                                + " (size=" + store.currentTemplate.length() + "B)");
                    } else {
                        System.out.println("[OO-CB] [WARN] 下载文档为空");
                    }
                    tmp.delete();
                } catch (Exception e) {
                    System.out.println("[OO-CB] [FAIL] 保存失败: " + e.getMessage());
                    return ResponseEntity.ok(Map.of("key", key, "error", 1));
                }
            }
        }
        return ResponseEntity.ok(Map.of("key", key, "error", 0));
    }

    /** 前端轮询：最新保存的 key + 时间戳 */
    @GetMapping("/api/oo/save_status")
    public Map<String, Object> ooSaveStatus() {
        return Map.of("saved_key", lastSaveKey == null ? "" : lastSaveKey, "saved_ts", lastSaveTs);
    }

    /** 前端注册当前活跃编辑器 key */
    @PostMapping("/api/oo/active_key")
    public ResponseEntity<?> ooActiveKey(@RequestBody(required = false) Map<String, Object> data) {
        if (data == null) data = Map.of();
        String key = data.get("key") == null ? null : String.valueOf(data.get("key"));
        if (key != null && !key.isEmpty()) {
            activeDocKey = key;
            System.out.println("[OO-AK] active key = " + key.substring(0, Math.min(28, key.length())));
        }
        return ResponseEntity.ok(Map.of("error", 0));
    }

    /** OO CommandService forcesave（命令名必须全小写 forcesave；error=4 表示无新修改视为成功） */
    @PostMapping("/api/oo/force_save")
    public ResponseEntity<?> ooForceSave(@RequestBody(required = false) Map<String, Object> data) {
        if (data == null) data = Map.of();
        String key = data.get("key") == null ? null : String.valueOf(data.get("key"));
        if (key == null || key.isEmpty()) return err("missing key", HttpStatus.BAD_REQUEST);
        try {
            HttpRequest req = HttpRequest.newBuilder(
                    URI.create(OO_PUBLIC_URL + "/coauthoring/CommandService.ashx?shardkey="
                            + java.net.URLEncoder.encode(key, StandardCharsets.UTF_8)))
                    .header("Content-Type", "application/json")
                    .POST(HttpRequest.BodyPublishers.ofString(
                            mapper.writeValueAsString(Map.of("c", "forcesave", "key", key))))
                    .build();
            HttpResponse<String> resp = http.send(req, HttpResponse.BodyHandlers.ofString());
            System.out.println("[OO-FS] CommandService 响应: HTTP " + resp.statusCode()
                    + " CT=" + resp.headers().firstValue("Content-Type").orElse("?")
                    + " body=" + resp.body().substring(0, Math.min(200, resp.body().length())));
            Map<String, Object> body = mapper.readValue(resp.body(), Map.class);
            System.out.println("[OO-FS] forcesave key=" + key.substring(0, Math.min(28, key.length()))
                    + " -> resp=" + body);
            if (Integer.valueOf(4).equals(((Number) body.getOrDefault("error", 1)).intValue())) {
                return ResponseEntity.ok(Map.of("error", 0, "no_changes", true, "resp", body));
            }
            return ResponseEntity.ok(Map.of("error", ((Number) body.getOrDefault("error", 1)).intValue(), "resp", body));
        } catch (Exception e) {
            System.out.println("[OO-FS] [FAIL] forcesave 失败: " + e.getMessage());
            Map<String, Object> resp = new LinkedHashMap<>();
            resp.put("error", String.valueOf(e.getMessage()));
            return ResponseEntity.ok(resp);
        }
    }

    /** OnlyOffice 静态托管（模板/标注 docx；vite 代理 /static-docx → 本路由） */
    @GetMapping("/static-docx/{fname}")
    public ResponseEntity<?> staticDocx(@PathVariable String fname) {
        File path = new File(store.staticDocsDir, new File(fname).getName());  // 防目录穿越
        if (!path.exists()) return err("文档不存在: " + fname, HttpStatus.NOT_FOUND);
        return fileResponse(path, fname, false);
    }
}
