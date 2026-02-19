import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CodeAnalysisTool {

    // 配置项：仅保留唯一根目录和Excel输出路径
    private static String SCAN_BASE_DIR;
    private static String EXCEL_OUTPUT_PATH;
    // 扫描根目录的File对象
    private static File ROOT_DIR;
    // 缓存：类名→文件相对路径（仅保留，不影响核心功能）
    private static Map<String, String> classNameToRelativePathMap = new HashMap<>();

    // ========== Struts配置重复检测相关 ==========
    private static Map<String, Integer> duplicateCountMap = new HashMap<>();
    private static final String DUPLICATE_KEY_SPLIT = "_";

    // ========== 前端文件.do路径提取相关 ==========
    private static final Pattern DO_PATH_PATTERN = Pattern.compile("([\"'\\s])(/[^\"'\\s]+\\.do)([\"'\\s])");
    private static List<Map<String, String>> frontEndDoPathList = new ArrayList<>();
    private static final List<String> FRONT_END_SUFFIX = Arrays.asList("jsp", "html", "js");

    // ========== Action类扫描相关（已调整，不区分大小写） ==========
    private static List<Map<String, String>> actionClassList = new ArrayList<>();
    // 匹配类定义：public class XxxAction extends Yyy（大小写不敏感）
    private static final Pattern CLASS_DEF_PATTERN = Pattern.compile("public\\s+class\\s+(\\w+)\\s*(extends\\s+(\\w+))?.*", Pattern.CASE_INSENSITIVE);

    // ========== Service接口扫描相关（不区分大小写） ==========
    private static List<Map<String, String>> serviceInterfaceList = new ArrayList<>();
    // 匹配public interface XxxService {（大小写不敏感）
    private static final Pattern INTERFACE_DEF_PATTERN = Pattern.compile("public\\s+interface\\s+(\\w+service)\\s*\\{?", Pattern.CASE_INSENSITIVE);
    // 去重：存储小写类名，避免重复（适配不规范命名）
    private static Set<String> serviceInterfaceSet = new HashSet<>();

    // ========== ServiceImpl实现类扫描相关（移除Manager字段，不区分大小写） ==========
    private static List<Map<String, String>> serviceImplList = new ArrayList<>();
    // 匹配public class XxxServiceImpl implements XXX（大小写不敏感）
    private static final Pattern SERVICE_IMPL_DEF_PATTERN = Pattern.compile("public\\s+class\\s+(\\w+serviceimpl)\\s*(implements\\s+([^\\{]+))?.*", Pattern.CASE_INSENSITIVE);
    private static final Pattern SERVICE_ANNOTATION_PATTERN = Pattern.compile("@service\\s*(\\(.*\\))?", Pattern.CASE_INSENSITIVE);
    private static final Pattern SOFA_SERVICE_ANNOTATION_PATTERN = Pattern.compile("@sofaservice\\s*\\(([^\\)]*)\\)", Pattern.CASE_INSENSITIVE);
    private static final Pattern BINDING_TYPE_PATTERN = Pattern.compile("bindingtype\\s*=\\s*[\"']([^\"']+)['\"]", Pattern.CASE_INSENSITIVE);
    private static Set<String> serviceImplSet = new HashSet<>();

    // ========== Manager接口扫描相关（不区分大小写） ==========
    private static List<Map<String, String>> managerInterfaceList = new ArrayList<>();
    // 匹配public interface XxxManager {（大小写不敏感）
    private static final Pattern MANAGER_INTERFACE_DEF_PATTERN = Pattern.compile("public\\s+interface\\s+(\\w+manager)\\s*\\{?", Pattern.CASE_INSENSITIVE);
    private static Set<String> managerInterfaceSet = new HashSet<>();

    // ========== ManagerImpl实现类扫描相关（移除Dao字段，不区分大小写） ==========
    private static List<Map<String, String>> managerImplList = new ArrayList<>();
    // 匹配public class XxxManagerImpl implements XXX（大小写不敏感）
    private static final Pattern MANAGER_IMPL_DEF_PATTERN = Pattern.compile("public\\s+class\\s+(\\w+managerimpl)\\s*(implements\\s+([^\\{]+))?.*", Pattern.CASE_INSENSITIVE);
    private static final Pattern TRANSACTIONAL_ANNOTATION_PATTERN = Pattern.compile("@transactional\\s*(\\(.*\\))?", Pattern.CASE_INSENSITIVE);
    private static Set<String> managerImplSet = new HashSet<>();

    // ========== Dao接口扫描相关（不区分大小写） ==========
    private static List<Map<String, String>> daoInterfaceList = new ArrayList<>();
    // 匹配public interface XxxDao {（大小写不敏感）
    private static final Pattern DAO_INTERFACE_DEF_PATTERN = Pattern.compile("public\\s+interface\\s+(\\w+dao)\\s*\\{?", Pattern.CASE_INSENSITIVE);
    private static Set<String> daoInterfaceSet = new HashSet<>();

    // ========== DaoImpl实现类扫描相关（不区分大小写） ==========
    private static List<Map<String, String>> daoImplList = new ArrayList<>();
    // 匹配public class XxxDaoImpl implements XXX（大小写不敏感）
    private static final Pattern DAO_IMPL_DEF_PATTERN = Pattern.compile("public\\s+class\\s+(\\w+daoimpl)\\s*(implements\\s+([^\\{]+))?.*", Pattern.CASE_INSENSITIVE);
    private static final Pattern REPOSITORY_ANNOTATION_PATTERN = Pattern.compile("@repository\\s*(\\(.*\\))?", Pattern.CASE_INSENSITIVE);
    private static Set<String> daoImplSet = new HashSet<>();

    public static void main(String[] args) {
        try {
            // 消除Log4j2报错
            System.setProperty("org.apache.logging.log4j.simplelog.defaultloglevel", "info");
            System.setProperty("log4j2.loggerContextFactory", "org.apache.logging.log4j.simple.SimpleLoggerContextFactory");

            // 1. 加载配置文件
            loadConfig();

            // 2. 校验配置
            if (SCAN_BASE_DIR == null || SCAN_BASE_DIR.isEmpty() ||
                    EXCEL_OUTPUT_PATH == null || EXCEL_OUTPUT_PATH.isEmpty()) {
                System.err.println("❌ 配置文件缺失必要项！需包含scan.base.dir、excel.output.path");
                return;
            }

            // 初始化唯一根目录
            ROOT_DIR = new File(SCAN_BASE_DIR);
            if (!ROOT_DIR.exists() || !ROOT_DIR.isDirectory()) {
                System.err.println("❌ 扫描/源码根目录不存在：" + SCAN_BASE_DIR);
                return;
            }

            // 预扫描所有Java文件
            System.out.println("========== 预扫描所有Java文件 ==========");
            preScanAllJavaFiles(ROOT_DIR);

            // ========== 3. 扫描解析Struts配置文件 ==========
            List<Map<String, String>> strutsConfigList = new ArrayList<>();
            scanAllStrutsConfig(ROOT_DIR, strutsConfigList);
            markDuplicateItems(strutsConfigList);
            printDuplicateSummary();

            // ========== 4. 扫描前端文件提取.do路径 ==========
            System.out.println("\n========== 扫描根目录下前端文件提取.do路径 ==========");
            scanFrontEndFiles(ROOT_DIR);
            countDoPathInFile();

            // ========== 5. 扫描所有以Action结尾的Java文件（不区分大小写） ==========
            System.out.println("\n========== 扫描所有以Action结尾的Java文件 ==========");
            scanAllActionJavaFiles(ROOT_DIR);

            // ========== 6. 扫描所有Java文件，提取以Service结尾的接口（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取Service接口 ==========");
            scanAllServiceInterfaces(ROOT_DIR);

            // ========== 7. 扫描所有Java文件，提取以ServiceImpl结尾的实现类（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取ServiceImpl实现类 ==========");
            scanAllServiceImplClasses(ROOT_DIR);

            // ========== 8. 扫描所有Java文件，提取以Manager结尾的接口（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取Manager接口 ==========");
            scanAllManagerInterfaces(ROOT_DIR);

            // ========== 9. 扫描所有Java文件，提取以ManagerImpl结尾的实现类（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取ManagerImpl实现类 ==========");
            scanAllManagerImplClasses(ROOT_DIR);

            // ========== 10. 扫描所有Java文件，提取以Dao结尾的接口（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取Dao接口 ==========");
            scanAllDaoInterfaces(ROOT_DIR);

            // ========== 11. 扫描所有Java文件，提取以DaoImpl结尾的实现类（不区分大小写） ==========
            System.out.println("\n========== 扫描所有Java文件提取DaoImpl实现类 ==========");
            scanAllDaoImplClasses(ROOT_DIR);

            // ========== 12. 写入9个Sheet的Excel ==========
            writeExcel(strutsConfigList, frontEndDoPathList, actionClassList,
                    serviceInterfaceList, serviceImplList, managerInterfaceList, managerImplList, daoInterfaceList, daoImplList, EXCEL_OUTPUT_PATH);

            System.out.println("\n✅ 全部解析完成！");
            System.out.println("   - Struts配置记录数：" + strutsConfigList.size());
            System.out.println("   - 前端DO路径记录数：" + frontEndDoPathList.size());
            System.out.println("   - Action类记录数：" + actionClassList.size());
            System.out.println("   - Service接口记录数：" + serviceInterfaceList.size());
            System.out.println("   - ServiceImpl实现类记录数：" + serviceImplList.size());
            System.out.println("   - Manager接口记录数：" + managerInterfaceList.size());
            System.out.println("   - ManagerImpl实现类记录数：" + managerImplList.size());
            System.out.println("   - Dao接口记录数：" + daoInterfaceList.size());
            System.out.println("   - DaoImpl实现类记录数：" + daoImplList.size());
            System.out.println("✅ Excel生成路径：" + EXCEL_OUTPUT_PATH);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 加载配置文件
     */
    private static void loadConfig() throws Exception {
        Properties props = new Properties();
        File configFile = new File("config.properties");
        InputStream is = configFile.exists() ? new FileInputStream(configFile) :
                StrutsConfigFullParser.class.getClassLoader().getResourceAsStream("config.properties");

        if (is == null) {
            throw new RuntimeException("❌ 未找到config.properties配置文件");
        }
        props.load(is);
        is.close();

        SCAN_BASE_DIR = props.getProperty("scan.base.dir").trim();
        EXCEL_OUTPUT_PATH = props.getProperty("excel.output.path").trim();
    }

    /**
     * 预扫描所有Java文件
     */
    private static void preScanAllJavaFiles(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                preScanAllJavaFiles(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    String content = readFileContent(file);
                    String className = extractClassName(content);
                    if (className != null && !className.isEmpty()) {
                        String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                        classNameToRelativePathMap.put(className, relativePath);
                    }
                }
            }
        }
    }

    /**
     * 提取Java文件的类名/接口名（适配不区分大小写）
     */
    private static String extractClassName(String content) {
        Matcher classMatcher = CLASS_DEF_PATTERN.matcher(content);
        if (classMatcher.find()) {
            return classMatcher.group(1).trim();
        }
        // 提取Service接口名（正则已大小写不敏感）
        Matcher interfaceMatcher = INTERFACE_DEF_PATTERN.matcher(content);
        if (interfaceMatcher.find()) {
            return interfaceMatcher.group(1).trim();
        }
        // 提取ServiceImpl类名（正则已大小写不敏感）
        Matcher serviceImplMatcher = SERVICE_IMPL_DEF_PATTERN.matcher(content);
        if (serviceImplMatcher.find()) {
            return serviceImplMatcher.group(1).trim();
        }
        // 提取Manager接口名（正则已大小写不敏感）
        Matcher managerInterfaceMatcher = MANAGER_INTERFACE_DEF_PATTERN.matcher(content);
        if (managerInterfaceMatcher.find()) {
            return managerInterfaceMatcher.group(1).trim();
        }
        // 提取ManagerImpl类名（正则已大小写不敏感）
        Matcher managerImplMatcher = MANAGER_IMPL_DEF_PATTERN.matcher(content);
        if (managerImplMatcher.find()) {
            return managerImplMatcher.group(1).trim();
        }
        // 提取Dao接口名（正则已大小写不敏感）
        Matcher daoInterfaceMatcher = DAO_INTERFACE_DEF_PATTERN.matcher(content);
        if (daoInterfaceMatcher.find()) {
            return daoInterfaceMatcher.group(1).trim();
        }
        // 提取DaoImpl类名（正则已大小写不敏感）
        Matcher daoImplMatcher = DAO_IMPL_DEF_PATTERN.matcher(content);
        if (daoImplMatcher.find()) {
            return daoImplMatcher.group(1).trim();
        }
        return null;
    }

    // ===================== Struts配置解析 =====================
    private static void scanAllStrutsConfig(File dir, List<Map<String, String>> resultList) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllStrutsConfig(file, resultList);
            } else {
                String fileName = file.getName().toLowerCase();
                if (fileName.contains("struts") && fileName.endsWith(".xml")) {
                    System.out.println("解析Struts配置：" + file.getAbsolutePath());
                    parseSingleStrutsConfig(file, resultList);
                }
            }
        }
    }

    private static void parseSingleStrutsConfig(File file, List<Map<String, String>> resultList) throws Exception {
        String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);

        SAXReader reader = new SAXReader();
        Document document = reader.read(file);
        Element root = document.getRootElement();

        Map<String, String> formBeanMap = new HashMap<>();
        List<Element> formBeans = root.selectNodes("//form-beans/form-bean");
        for (Element bean : formBeans) {
            formBeanMap.put(bean.attributeValue("name"), bean.attributeValue("type"));
        }

        List<Element> actionList = root.selectNodes("//action-mappings/action");
        for (Element actionEle : actionList) {
            String actionPath = actionEle.attributeValue("path");
            String actionType = actionEle.attributeValue("type");
            String actionName = actionEle.attributeValue("name");
            String formBeanType = formBeanMap.getOrDefault(actionName, "");

            List<Element> forwards = actionEle.selectNodes("forward");
            if (forwards == null || forwards.isEmpty()) {
                String duplicateKey = getDuplicateKey(actionPath, formBeanType, "");
                duplicateCountMap.put(duplicateKey, duplicateCountMap.getOrDefault(duplicateKey, 0) + 1);
                addStrutsRow(resultList, relativePath, formBeanType, actionPath, actionType, actionName, "", "", duplicateKey);
            } else {
                for (Element forward : forwards) {
                    String forwardName = forward.attributeValue("name");
                    String forwardPath = forward.attributeValue("path");
                    String duplicateKey = getDuplicateKey(actionPath, formBeanType, forwardName);
                    duplicateCountMap.put(duplicateKey, duplicateCountMap.getOrDefault(duplicateKey, 0) + 1);
                    addStrutsRow(resultList, relativePath, formBeanType, actionPath, actionType, actionName, forwardName, forwardPath, duplicateKey);
                }
            }
        }
    }

    private static String getDuplicateKey(String actionPath, String formBeanType, String forwardName) {
        actionPath = actionPath == null ? "空" : actionPath.trim();
        formBeanType = formBeanType == null ? "空" : formBeanType.trim();
        forwardName = forwardName == null ? "空" : forwardName.trim();
        return actionPath + DUPLICATE_KEY_SPLIT + formBeanType + DUPLICATE_KEY_SPLIT + forwardName;
    }

    private static void addStrutsRow(List<Map<String, String>> resultList,
                                     String relativePath,
                                     String formBean,
                                     String actionPath,
                                     String actionType,
                                     String actionName,
                                     String forwardName,
                                     String forwardPath,
                                     String duplicateKey) {
        Map<String, String> map = new HashMap<>();
        map.put("relativePath", relativePath);
        map.put("formBean", formBean);
        map.put("actionPath", actionPath);
        map.put("actionType", actionType);
        map.put("actionName", actionName);
        map.put("forwardName", forwardName);
        map.put("forwardPath", forwardPath);
        map.put("duplicateKey", duplicateKey);
        map.put("isDuplicate", "");
        map.put("duplicateCount", "");
        resultList.add(map);
    }

    private static void markDuplicateItems(List<Map<String, String>> resultList) {
        for (Map<String, String> record : resultList) {
            int count = duplicateCountMap.getOrDefault(record.get("duplicateKey"), 1);
            record.put("isDuplicate", count > 1 ? "是" : "否");
            record.put("duplicateCount", String.valueOf(count));
        }
    }

    private static void printDuplicateSummary() {
        System.out.println("\n========== Struts配置重复项汇总 ==========");
        int totalDuplicate = 0;
        for (Map.Entry<String, Integer> entry : duplicateCountMap.entrySet()) {
            if (entry.getValue() > 1) {
                totalDuplicate++;
                String[] parts = entry.getKey().split(DUPLICATE_KEY_SPLIT);
                System.out.println(String.format(
                        "重复：action-path=%s, form-bean=%s, forward=%s | 次数：%d",
                        parts[0], parts[1], parts[2], entry.getValue()
                ));
            }
        }
        System.out.println(totalDuplicate == 0 ? "✅ 无重复Struts配置" : "⚠️  共" + totalDuplicate + "组重复配置");
    }

    // ===================== 前端文件DO路径提取 =====================
    private static void scanFrontEndFiles(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanFrontEndFiles(file);
            } else {
                String fileName = file.getName().toLowerCase();
                String suffix = fileName.lastIndexOf(".") > 0 ? fileName.substring(fileName.lastIndexOf(".") + 1) : "";
                if (FRONT_END_SUFFIX.contains(suffix)) {
                    System.out.println("解析前端文件：" + file.getAbsolutePath());
                    extractDoPathFromFile(file);
                }
            }
        }
    }

    private static void extractDoPathFromFile(File file) throws Exception {
        String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
        String content = readFileContent(file);

        Matcher matcher = DO_PATH_PATTERN.matcher(content);
        Set<String> doPathSet = new HashSet<>();
        while (matcher.find()) {
            String doPath = matcher.group(2).trim();
            if (!doPath.isEmpty()) {
                doPathSet.add(doPath);
            }
        }

        for (String doPath : doPathSet) {
            Map<String, String> map = new HashMap<>();
            map.put("fileRelativePath", relativePath);
            map.put("doPath", doPath);
            map.put("count", "1");
            frontEndDoPathList.add(map);
        }
    }

    private static String readFileContent(File file) throws Exception {
        try (java.io.BufferedReader br = new java.io.BufferedReader(
                new java.io.InputStreamReader(new FileInputStream(file), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = br.readLine()) != null) {
                sb.append(line).append("\n");
            }
            return sb.toString();
        } catch (Exception e) {
            try (java.io.BufferedReader br = new java.io.BufferedReader(
                    new java.io.InputStreamReader(new FileInputStream(file), "GBK"))) {
                StringBuilder sb = new StringBuilder();
                String line;
                while ((line = br.readLine()) != null) {
                    sb.append(line).append("\n");
                }
                return sb.toString();
            }
        }
    }

    private static void countDoPathInFile() {
        Map<String, Integer> tempCountMap = new HashMap<>();
        for (Map<String, String> record : frontEndDoPathList) {
            String key = record.get("fileRelativePath") + "_" + record.get("doPath");
            tempCountMap.put(key, tempCountMap.getOrDefault(key, 0) + 1);
        }

        for (Map<String, String> record : frontEndDoPathList) {
            String key = record.get("fileRelativePath") + "_" + record.get("doPath");
            record.put("count", String.valueOf(tempCountMap.get(key)));
        }
    }

    // ===================== 扫描以Action结尾的Java文件（不区分大小写） =====================
    private static void scanAllActionJavaFiles(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllActionJavaFiles(file);
            } else {
                String fileName = file.getName();
                // 筛选：.java文件 + 文件名（小写）以action结尾（不区分大小写）
                if (fileName.endsWith(".java")) {
                    String nameWithoutExt = fileName.substring(0, fileName.lastIndexOf("."));
                    if (nameWithoutExt.toLowerCase().endsWith("action")) {
                        System.out.println("扫描Action文件：" + file.getAbsolutePath());
                        parseActionJavaFile(file);
                    }
                }
            }
        }
    }

    private static void parseActionJavaFile(File file) throws Exception {
        // 1. 提取Action类名
        String content = readFileContent(file);
        String actionClassName = extractClassName(content);
        if (actionClassName == null || actionClassName.isEmpty()) {
            actionClassName = "未提取到类名";
        }

        // 2. 提取包名
        String packageName = extractPackageName(content);

        // 3. 提取文件相对路径
        String actionRelativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);

        // 4. 提取父类名称（仅名称，不查相对路径）
        String parentClassName = extractParentClassName(content);
        if (parentClassName == null || parentClassName.isEmpty()) {
            parentClassName = "无父类";
        }

        // 5. 组装数据
        addActionClassRow(actionClassName, packageName, actionRelativePath, parentClassName);
    }

    private static String extractParentClassName(String content) {
        Matcher classMatcher = CLASS_DEF_PATTERN.matcher(content);
        if (classMatcher.find()) {
            String parentName = classMatcher.group(3);
            return parentName == null || parentName.isEmpty() ? "" : parentName.trim();
        }
        return "";
    }

    private static String extractPackageName(String content) {
        // 包名匹配不区分大小写
        Pattern packagePattern = Pattern.compile("package\\s+([^;]+);", Pattern.CASE_INSENSITIVE);
        Matcher packageMatcher = packagePattern.matcher(content);
        return packageMatcher.find() ? packageMatcher.group(1).trim() : "无包名";
    }

    private static void addActionClassRow(String actionClassName, String packageName, String actionRelativePath, String parentClassName) {
        Map<String, String> map = new HashMap<>();
        map.put("actionClassName", actionClassName);    // Action类名
        map.put("packageName", packageName);            // 包名
        map.put("actionRelativePath", actionRelativePath); // 文件相对路径
        map.put("parentClassName", parentClassName);    // 父类名称
        actionClassList.add(map);
    }

    // ===================== 扫描所有Service接口（不区分大小写） =====================
    private static void scanAllServiceInterfaces(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllServiceInterfaces(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseServiceInterfaceFile(file);
                }
            }
        }
    }

    private static void parseServiceInterfaceFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher interfaceMatcher = INTERFACE_DEF_PATTERN.matcher(content);

        while (interfaceMatcher.find()) {
            String interfaceName = interfaceMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxSERVICE、Xxxservice等不规范命名）
            String interfaceNameLower = interfaceName.toLowerCase();
            if (!serviceInterfaceSet.contains(interfaceNameLower)) {
                serviceInterfaceSet.add(interfaceNameLower);
                // 接口文件相对路径
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                // 提取包名
                String packageName = extractPackageName(content);

                // 组装数据
                addServiceInterfaceRow(interfaceName, packageName, relativePath);
                System.out.println("找到Service接口：" + interfaceName + " → " + relativePath);
            }
        }
    }

    private static void addServiceInterfaceRow(String interfaceName, String packageName, String relativePath) {
        Map<String, String> map = new HashMap<>();
        map.put("interfaceName", interfaceName);       // 接口名（保留原大小写）
        map.put("packageName", packageName);           // 包名
        map.put("fileRelativePath", relativePath);     // 接口文件相对路径
        serviceInterfaceList.add(map);
    }

    // ===================== 扫描所有ServiceImpl实现类（移除Manager字段，不区分大小写） =====================
    private static void scanAllServiceImplClasses(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllServiceImplClasses(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseServiceImplFile(file);
                }
            }
        }
    }

    private static void parseServiceImplFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher serviceImplMatcher = SERVICE_IMPL_DEF_PATTERN.matcher(content);

        while (serviceImplMatcher.find()) {
            String implClassName = serviceImplMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxSERVICEIMPL、Xxxserviceimpl等）
            String implClassNameLower = implClassName.toLowerCase();
            if (!serviceImplSet.contains(implClassNameLower)) {
                serviceImplSet.add(implClassNameLower);

                // 1. 提取被实现的接口名
                String implementedInterfaces = extractImplementedInterfaces(serviceImplMatcher.group(3));

                // 2. 检查是否有@Service注解（大小写不敏感）
                boolean hasServiceAnnotation = SERVICE_ANNOTATION_PATTERN.matcher(content).find();

                // 3. 检查是否有@SofaService注解，并提取bindingType（大小写不敏感）
                String sofaServiceInfo = extractSofaServiceInfo(content);
                boolean hasSofaServiceAnnotation = !sofaServiceInfo.startsWith("无");
                String bindingType = sofaServiceInfo.split("\\|")[1];

                // 4. 基础信息
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                String packageName = extractPackageName(content);

                // 组装数据（移除Manager字段）
                addServiceImplRow(
                        implClassName, packageName, relativePath,
                        implementedInterfaces,
                        hasServiceAnnotation ? "是" : "否",
                        hasSofaServiceAnnotation ? "是" : "否",
                        bindingType
                );

                System.out.println("找到ServiceImpl类：" + implClassName + " → " + relativePath);
            }
        }
    }

    private static String extractImplementedInterfaces(String implementsPart) {
        if (implementsPart == null || implementsPart.trim().isEmpty()) {
            return "无实现接口";
        }

        List<String> interfaceNames = new ArrayList<>();
        // 移除泛型内容（如<User>）
        String cleanPart = implementsPart.replaceAll("<[^>]+>", "");
        // 分割多个接口（支持,分割）
        String[] parts = cleanPart.split(",");

        for (String part : parts) {
            String interfaceName = part.trim();
            if (!interfaceName.isEmpty() && !interfaceName.equals("")) {
                interfaceNames.add(interfaceName);
            }
        }

        return interfaceNames.isEmpty() ? "无实现接口" : String.join(", ", interfaceNames);
    }

    private static String extractSofaServiceInfo(String content) {
        Matcher sofaServiceMatcher = SOFA_SERVICE_ANNOTATION_PATTERN.matcher(content);
        if (!sofaServiceMatcher.find()) {
            return "无|无";
        }

        String annotationContent = sofaServiceMatcher.group(1);
        Matcher bindingTypeMatcher = BINDING_TYPE_PATTERN.matcher(annotationContent);
        String bindingType = bindingTypeMatcher.find() ? bindingTypeMatcher.group(1).trim() : "未指定";

        return "有|" + bindingType;
    }

    private static void addServiceImplRow(
            String implClassName, String packageName, String relativePath,
            String implementedInterfaces, String hasServiceAnnotation,
            String hasSofaServiceAnnotation, String bindingType
    ) {
        Map<String, String> map = new HashMap<>();
        map.put("implClassName", implClassName);               // ServiceImpl类名（保留原大小写）
        map.put("packageName", packageName);                   // 包名
        map.put("fileRelativePath", relativePath);             // 文件相对路径
        map.put("implementedInterfaces", implementedInterfaces); // 被实现的接口名
        map.put("hasServiceAnnotation", hasServiceAnnotation); // 是否有@Service注解
        map.put("hasSofaServiceAnnotation", hasSofaServiceAnnotation); // 是否有@SofaService注解
        map.put("bindingType", bindingType);                   // bindingType值
        serviceImplList.add(map);
    }

    // ===================== 扫描所有Manager接口（不区分大小写） =====================
    private static void scanAllManagerInterfaces(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllManagerInterfaces(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseManagerInterfaceFile(file);
                }
            }
        }
    }

    private static void parseManagerInterfaceFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher managerInterfaceMatcher = MANAGER_INTERFACE_DEF_PATTERN.matcher(content);

        while (managerInterfaceMatcher.find()) {
            String interfaceName = managerInterfaceMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxMANAGER、Xxxmanager等）
            String interfaceNameLower = interfaceName.toLowerCase();
            if (!managerInterfaceSet.contains(interfaceNameLower)) {
                managerInterfaceSet.add(interfaceNameLower);
                // 接口文件相对路径
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                // 提取包名
                String packageName = extractPackageName(content);

                // 组装数据
                addManagerInterfaceRow(interfaceName, packageName, relativePath);
                System.out.println("找到Manager接口：" + interfaceName + " → " + relativePath);
            }
        }
    }

    private static void addManagerInterfaceRow(String interfaceName, String packageName, String relativePath) {
        Map<String, String> map = new HashMap<>();
        map.put("interfaceName", interfaceName);       // 接口名（保留原大小写）
        map.put("packageName", packageName);           // 包名
        map.put("fileRelativePath", relativePath);     // 接口文件相对路径
        managerInterfaceList.add(map);
    }

    // ===================== 扫描所有ManagerImpl实现类（移除Dao字段，不区分大小写） =====================
    private static void scanAllManagerImplClasses(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllManagerImplClasses(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseManagerImplFile(file);
                }
            }
        }
    }

    private static void parseManagerImplFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher managerImplMatcher = MANAGER_IMPL_DEF_PATTERN.matcher(content);

        while (managerImplMatcher.find()) {
            String implClassName = managerImplMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxMANAGERIMPL、Xxxmanagerimpl等）
            String implClassNameLower = implClassName.toLowerCase();
            if (!managerImplSet.contains(implClassNameLower)) {
                managerImplSet.add(implClassNameLower);

                // 1. 提取被实现的接口名
                String implementedInterfaces = extractImplementedInterfaces(managerImplMatcher.group(3));

                // 2. 检查是否有@Service注解（大小写不敏感）
                boolean hasServiceAnnotation = SERVICE_ANNOTATION_PATTERN.matcher(content).find();

                // 3. 检查是否有@Transactional注解（大小写不敏感）
                boolean hasTransactionalAnnotation = TRANSACTIONAL_ANNOTATION_PATTERN.matcher(content).find();

                // 4. 基础信息
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                String packageName = extractPackageName(content);

                // 组装数据（移除Dao字段）
                addManagerImplRow(
                        implClassName, packageName, relativePath,
                        implementedInterfaces,
                        hasServiceAnnotation ? "是" : "否",
                        hasTransactionalAnnotation ? "是" : "否"
                );

                System.out.println("找到ManagerImpl类：" + implClassName + " → " + relativePath);
            }
        }
    }

    private static void addManagerImplRow(
            String implClassName, String packageName, String relativePath,
            String implementedInterfaces, String hasServiceAnnotation,
            String hasTransactionalAnnotation
    ) {
        Map<String, String> map = new HashMap<>();
        map.put("implClassName", implClassName);               // ManagerImpl类名（保留原大小写）
        map.put("packageName", packageName);                   // 包名
        map.put("fileRelativePath", relativePath);             // 文件相对路径
        map.put("implementedInterfaces", implementedInterfaces); // 被实现的接口名
        map.put("hasServiceAnnotation", hasServiceAnnotation); // 是否有@Service注解
        map.put("hasTransactionalAnnotation", hasTransactionalAnnotation); // 是否有@Transactional注解
        managerImplList.add(map);
    }

    // ===================== 扫描所有Dao接口（不区分大小写） =====================
    private static void scanAllDaoInterfaces(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllDaoInterfaces(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseDaoInterfaceFile(file);
                }
            }
        }
    }

    private static void parseDaoInterfaceFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher daoInterfaceMatcher = DAO_INTERFACE_DEF_PATTERN.matcher(content);

        while (daoInterfaceMatcher.find()) {
            String interfaceName = daoInterfaceMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxDAO、Xxxdao等）
            String interfaceNameLower = interfaceName.toLowerCase();
            if (!daoInterfaceSet.contains(interfaceNameLower)) {
                daoInterfaceSet.add(interfaceNameLower);
                // 接口文件相对路径
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                // 提取包名
                String packageName = extractPackageName(content);

                // 组装数据
                addDaoInterfaceRow(interfaceName, packageName, relativePath);
                System.out.println("找到Dao接口：" + interfaceName + " → " + relativePath);
            }
        }
    }

    private static void addDaoInterfaceRow(String interfaceName, String packageName, String relativePath) {
        Map<String, String> map = new HashMap<>();
        map.put("interfaceName", interfaceName);       // 接口名（保留原大小写）
        map.put("packageName", packageName);           // 包名
        map.put("fileRelativePath", relativePath);     // 接口文件相对路径
        daoInterfaceList.add(map);
    }

    // ===================== 扫描所有DaoImpl实现类（不区分大小写） =====================
    private static void scanAllDaoImplClasses(File dir) throws Exception {
        if (!dir.isDirectory() || dir.getName().equalsIgnoreCase("build")) {
            return;
        }

        File[] files = dir.listFiles();
        if (files == null) return;

        for (File file : files) {
            if (file.isDirectory()) {
                scanAllDaoImplClasses(file);
            } else {
                String fileName = file.getName();
                if (fileName.endsWith(".java")) {
                    parseDaoImplFile(file);
                }
            }
        }
    }

    private static void parseDaoImplFile(File file) throws Exception {
        String content = readFileContent(file);
        Matcher daoImplMatcher = DAO_IMPL_DEF_PATTERN.matcher(content);

        while (daoImplMatcher.find()) {
            String implClassName = daoImplMatcher.group(1).trim();
            // 去重：存储小写类名，避免重复（适配XxxDAOIMPL、Xxxdaoimpl等）
            String implClassNameLower = implClassName.toLowerCase();
            if (!daoImplSet.contains(implClassNameLower)) {
                daoImplSet.add(implClassNameLower);

                // 1. 提取被实现的接口名
                String implementedInterfaces = extractImplementedInterfaces(daoImplMatcher.group(3));

                // 2. 检查是否有@Repository注解（大小写不敏感）
                boolean hasRepositoryAnnotation = REPOSITORY_ANNOTATION_PATTERN.matcher(content).find();

                // 3. 基础信息
                String relativePath = ROOT_DIR.toURI().relativize(file.toURI()).getPath().replace("/", File.separator);
                String packageName = extractPackageName(content);

                // 组装数据
                addDaoImplRow(
                        implClassName, packageName, relativePath,
                        implementedInterfaces,
                        hasRepositoryAnnotation ? "是" : "否"
                );

                System.out.println("找到DaoImpl类：" + implClassName + " → " + relativePath);
            }
        }
    }

    private static void addDaoImplRow(
            String implClassName, String packageName, String relativePath,
            String implementedInterfaces, String hasRepositoryAnnotation
    ) {
        Map<String, String> map = new HashMap<>();
        map.put("implClassName", implClassName);               // DaoImpl类名（保留原大小写）
        map.put("packageName", packageName);                   // 包名
        map.put("fileRelativePath", relativePath);             // 文件相对路径
        map.put("implementedInterfaces", implementedInterfaces); // 被实现的接口名
        map.put("hasRepositoryAnnotation", hasRepositoryAnnotation); // 是否有@Repository注解
        daoImplList.add(map);
    }

    // ===================== Excel写入 =====================
    private static void writeExcel(List<Map<String, String>> strutsList,
                                   List<Map<String, String>> frontEndList,
                                   List<Map<String, String>> actionClassList,
                                   List<Map<String, String>> serviceInterfaceList,
                                   List<Map<String, String>> serviceImplList,
                                   List<Map<String, String>> managerInterfaceList,
                                   List<Map<String, String>> managerImplList,
                                   List<Map<String, String>> daoInterfaceList,
                                   List<Map<String, String>> daoImplList,
                                   String outPath) throws Exception {
        Workbook workbook = new XSSFWorkbook();

        // ========== Sheet1：Struts配置 ==========
        Sheet sheet1 = workbook.createSheet("Struts配置");
        String[] strutsHeaders = {
                "配置文件相对路径", "form-bean-type", "action-path", "action-type",
                "action-name(form)", "forward-name", "forward-path", "是否重复", "重复次数"
        };
        Row sheet1Head = sheet1.createRow(0);
        for (int i = 0; i < strutsHeaders.length; i++) {
            sheet1Head.createCell(i).setCellValue(strutsHeaders[i]);
        }
        int row1 = 1;
        for (Map<String, String> map : strutsList) {
            Row row = sheet1.createRow(row1++);
            row.createCell(0).setCellValue(map.get("relativePath"));
            row.createCell(1).setCellValue(map.get("formBean"));
            row.createCell(2).setCellValue(map.get("actionPath"));
            row.createCell(3).setCellValue(map.get("actionType"));
            row.createCell(4).setCellValue(map.get("actionName"));
            row.createCell(5).setCellValue(map.get("forwardName"));
            row.createCell(6).setCellValue(map.get("forwardPath"));
            row.createCell(7).setCellValue(map.get("isDuplicate"));
            row.createCell(8).setCellValue(map.get("duplicateCount"));
        }

        // ========== Sheet2：前端文件DO路径 ==========
        Sheet sheet2 = workbook.createSheet("前端文件DO路径");
        String[] frontEndHeaders = {"文件相对路径", ".do路径", "出现次数"};
        Row sheet2Head = sheet2.createRow(0);
        for (int i = 0; i < frontEndHeaders.length; i++) {
            sheet2Head.createCell(i).setCellValue(frontEndHeaders[i]);
        }
        int row2 = 1;
        for (Map<String, String> map : frontEndList) {
            Row row = sheet2.createRow(row2++);
            row.createCell(0).setCellValue(map.get("fileRelativePath"));
            row.createCell(1).setCellValue(map.get("doPath"));
            row.createCell(2).setCellValue(map.get("count"));
        }

        // ========== Sheet3：Action类信息 ==========
        Sheet sheet3 = workbook.createSheet("Action类信息");
        String[] actionClassHeaders = {
                "Action类名", "包名", "文件相对路径", "父类名称"
        };
        Row sheet3Head = sheet3.createRow(0);
        for (int i = 0; i < actionClassHeaders.length; i++) {
            sheet3Head.createCell(i).setCellValue(actionClassHeaders[i]);
        }
        int row3 = 1;
        // 排序：按Action类名（小写）升序排列，适配不规范命名
        Collections.sort(actionClassList, (o1, o2) -> o1.get("actionClassName").toLowerCase().compareTo(o2.get("actionClassName").toLowerCase()));
        for (Map<String, String> map : actionClassList) {
            Row row = sheet3.createRow(row3++);
            row.createCell(0).setCellValue(map.get("actionClassName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("actionRelativePath"));
            row.createCell(3).setCellValue(map.get("parentClassName"));
        }

        // ========== Sheet4：Service接口列表 ==========
        Sheet sheet4 = workbook.createSheet("Service接口列表");
        String[] serviceInterfaceHeaders = {
                "接口名", "包名", "接口文件相对路径"
        };
        Row sheet4Head = sheet4.createRow(0);
        for (int i = 0; i < serviceInterfaceHeaders.length; i++) {
            sheet4Head.createCell(i).setCellValue(serviceInterfaceHeaders[i]);
        }
        int row4 = 1;
        // 排序：按接口名（小写）升序排列
        Collections.sort(serviceInterfaceList, (o1, o2) -> o1.get("interfaceName").toLowerCase().compareTo(o2.get("interfaceName").toLowerCase()));
        for (Map<String, String> map : serviceInterfaceList) {
            Row row = sheet4.createRow(row4++);
            row.createCell(0).setCellValue(map.get("interfaceName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
        }

        // ========== Sheet5：ServiceImpl实现类列表（移除Manager字段列） ==========
        Sheet sheet5 = workbook.createSheet("ServiceImpl实现类列表");
        String[] serviceImplHeaders = {
                "ServiceImpl类名", "包名", "文件相对路径",
                "被实现的接口名", "是否有@Service注解", "是否有@SofaService注解",
                "bindingType值"
        };
        Row sheet5Head = sheet5.createRow(0);
        for (int i = 0; i < serviceImplHeaders.length; i++) {
            sheet5Head.createCell(i).setCellValue(serviceImplHeaders[i]);
        }
        int row5 = 1;
        // 排序：按ServiceImpl类名（小写）升序排列
        Collections.sort(serviceImplList, (o1, o2) -> o1.get("implClassName").toLowerCase().compareTo(o2.get("implClassName").toLowerCase()));
        for (Map<String, String> map : serviceImplList) {
            Row row = sheet5.createRow(row5++);
            row.createCell(0).setCellValue(map.get("implClassName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
            row.createCell(3).setCellValue(map.get("implementedInterfaces"));
            row.createCell(4).setCellValue(map.get("hasServiceAnnotation"));
            row.createCell(5).setCellValue(map.get("hasSofaServiceAnnotation"));
            row.createCell(6).setCellValue(map.get("bindingType"));
        }

        // ========== Sheet6：Manager接口列表 ==========
        Sheet sheet6 = workbook.createSheet("Manager接口列表");
        String[] managerInterfaceHeaders = {
                "接口名", "包名", "接口文件相对路径"
        };
        Row sheet6Head = sheet6.createRow(0);
        for (int i = 0; i < managerInterfaceHeaders.length; i++) {
            sheet6Head.createCell(i).setCellValue(managerInterfaceHeaders[i]);
        }
        int row6 = 1;
        // 排序：按接口名（小写）升序排列
        Collections.sort(managerInterfaceList, (o1, o2) -> o1.get("interfaceName").toLowerCase().compareTo(o2.get("interfaceName").toLowerCase()));
        for (Map<String, String> map : managerInterfaceList) {
            Row row = sheet6.createRow(row6++);
            row.createCell(0).setCellValue(map.get("interfaceName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
        }

        // ========== Sheet7：ManagerImpl实现类列表（移除Dao字段列） ==========
        Sheet sheet7 = workbook.createSheet("ManagerImpl实现类列表");
        String[] managerImplHeaders = {
                "ManagerImpl类名", "包名", "文件相对路径",
                "被实现的接口名", "是否有@Service注解", "是否有@Transactional注解"
        };
        Row sheet7Head = sheet7.createRow(0);
        for (int i = 0; i < managerImplHeaders.length; i++) {
            sheet7Head.createCell(i).setCellValue(managerImplHeaders[i]);
        }
        int row7 = 1;
        // 排序：按ManagerImpl类名（小写）升序排列
        Collections.sort(managerImplList, (o1, o2) -> o1.get("implClassName").toLowerCase().compareTo(o2.get("implClassName").toLowerCase()));
        for (Map<String, String> map : managerImplList) {
            Row row = sheet7.createRow(row7++);
            row.createCell(0).setCellValue(map.get("implClassName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
            row.createCell(3).setCellValue(map.get("implementedInterfaces"));
            row.createCell(4).setCellValue(map.get("hasServiceAnnotation"));
            row.createCell(5).setCellValue(map.get("hasTransactionalAnnotation"));
        }

        // ========== Sheet8：Dao接口列表 ==========
        Sheet sheet8 = workbook.createSheet("Dao接口列表");
        String[] daoInterfaceHeaders = {
                "接口名", "包名", "接口文件相对路径"
        };
        Row sheet8Head = sheet8.createRow(0);
        for (int i = 0; i < daoInterfaceHeaders.length; i++) {
            sheet8Head.createCell(i).setCellValue(daoInterfaceHeaders[i]);
        }
        int row8 = 1;
        // 排序：按接口名（小写）升序排列
        Collections.sort(daoInterfaceList, (o1, o2) -> o1.get("interfaceName").toLowerCase().compareTo(o2.get("interfaceName").toLowerCase()));
        for (Map<String, String> map : daoInterfaceList) {
            Row row = sheet8.createRow(row8++);
            row.createCell(0).setCellValue(map.get("interfaceName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
        }

        // ========== Sheet9：DaoImpl实现类列表 ==========
        Sheet sheet9 = workbook.createSheet("DaoImpl实现类列表");
        String[] daoImplHeaders = {
                "DaoImpl类名", "包名", "文件相对路径",
                "被实现的接口名", "是否有@Repository注解"
        };
        Row sheet9Head = sheet9.createRow(0);
        for (int i = 0; i < daoImplHeaders.length; i++) {
            sheet9Head.createCell(i).setCellValue(daoImplHeaders[i]);
        }
        int row9 = 1;
        // 排序：按DaoImpl类名（小写）升序排列
        Collections.sort(daoImplList, (o1, o2) -> o1.get("implClassName").toLowerCase().compareTo(o2.get("implClassName").toLowerCase()));
        for (Map<String, String> map : daoImplList) {
            Row row = sheet9.createRow(row9++);
            row.createCell(0).setCellValue(map.get("implClassName"));
            row.createCell(1).setCellValue(map.get("packageName"));
            row.createCell(2).setCellValue(map.get("fileRelativePath"));
            row.createCell(3).setCellValue(map.get("implementedInterfaces"));
            row.createCell(4).setCellValue(map.get("hasRepositoryAnnotation"));
        }

        // 自动调整列宽
        for (int i = 1; i <= 9; i++) {
            Sheet sheet = workbook.getSheetAt(i - 1);
            for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
                sheet.autoSizeColumn(j);
            }
        }

        // 写入文件
        try (FileOutputStream outputStream = new FileOutputStream(outPath)) {
            workbook.write(outputStream);
        }
        workbook.close();
    }
}
