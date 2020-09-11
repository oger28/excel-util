# excel-util
Excel导出工具类

 - 基于SpringBoot、POI、Swagger  
    
 * 设计思想：
 1. 样式最简化原则
 2. 约定大于规定原则

 
 * 优点：
 1. 可任意创建多个sheet
 2. 可在任意位置创建table
 3. 可按指定顺序导出实体类任意字段
    - 同一字段在不同table中名称可不同
    - 可通过Map<String,String> 的方式传入你想导出的字段
    - 可通过String[] headNames 和 String[] fieldNames 搭配的方式传入你想导出的字段
 4. 可一次性导出单sheet单表模式的Excel
 5. 实现二级树形表头合并的方式创建表
 6. 实现多级表头合并的方式创建表(兼容二级表头合并的方式)
 7. 动态设置列宽
 8. Controller 每一个接口对应一个demo
 
```
   /**
        * 单sheet单表模式，调用一个方法即可实现导出
        */
       @GetMapping("/exportStudents")
       @ApiOperation(value = "导出学生")
       @ResponseBody
       public void exportStudents(HttpServletResponse response) {
           String fileName = "学生列表";
           Map<String, String> headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           headMap.put("birthday", "生日");
           ExcelExportUtil.exportExcel(fileName, headMap, getStudents(), response);
       }
   
       /**
        * 多sheet 导出
        */
       @GetMapping("/exportStudentsAndTeachers")
       @ApiOperation(value = "导出学生和老师")
       @ResponseBody
       public void exportStudentsAndTeachers(HttpServletResponse response) {
           HSSFWorkbook workbook = new HSSFWorkbook();
           String fileName = "学生和老师列表";
   
           Map<String, String> headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           headMap.put("birthday", "生日");
           String sheetName = "学生列表";
           String tableName = "学生列表";
           Sheet sheet1 = workbook.createSheet(sheetName);
           ExcelExportUtil.createTable(tableName, headMap, getStudents(), sheet1, workbook);
   
           sheetName = "教师列表";
           tableName = "教师列表";
           headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           headMap.put("subject", "科目");
           Sheet sheet2 = workbook.createSheet(sheetName);
           ExcelExportUtil.createTable(tableName, headMap, getTeachers(), sheet2, workbook);
   
           ExcelExportUtil.exportExcel(fileName, workbook, response);
       }
   
       /**
        * 在一个 sheet 中有多张表
        */
       @GetMapping("/exportStudentsAndTeachers2")
       @ApiOperation(value = "导出学生和老师2")
       @ResponseBody
       public void exportStudentsAndTeachers2(HttpServletResponse response) {
           HSSFWorkbook workbook = new HSSFWorkbook();
           String fileName = "学生和老师列表";
           String sheetName = "学生和老师列表";
           Sheet sheet = workbook.createSheet(sheetName);
           Map<String, String> headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           headMap.put("birthday", "生日");
           String tableName = "学生列表";
           int line = ExcelExportUtil.createSheetTitle(fileName, sheet, workbook);
           line = ExcelExportUtil.createTable(line, tableName, headMap, getStudents(), sheet, workbook);
   
           tableName = "教师列表";
           headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           headMap.put("subject", "科目");
           ExcelExportUtil.createTable(line, tableName, headMap, getTeachers(), sheet, workbook);
   
           ExcelExportUtil.exportExcel(fileName, workbook, response);
       }
   
       /**
        * 二级合并表头单元格
        */
       @GetMapping("/exportStudentScores")
       @ApiOperation(value = "导出学生成绩表")
       @ResponseBody
       public void exportStudentScores(HttpServletResponse response) {
           String fileName = "学生成绩表";
           Map<String, Map<String, String>> mergeHeadMap = getMergeHeadMap();
           List<Student> scores = getScores();
           ExcelExportUtil.exportMergeHeadExcel(fileName, mergeHeadMap, scores, response);
       }
       
       /**
        * 多级合并表头单元格
        */
       @GetMapping("/exportStudentScores2")
       @ApiOperation(value = "导出学生成绩表")
       @ResponseBody
       public void exportStudentScores2(HttpServletResponse response) {
           String fileName = "学生成绩表";
           List<Map<String, Object>> mergeHeads = getMergeHeads();
           List<Student> scores = getScores();
           ExcelExportUtil.exportMergeHeadExcel(fileName, mergeHeads, scores, response);
       }
   
       private List<Map<String, Object>> getMergeHeads() {
           List<Map<String, Object>> mergeHeads = new ArrayList<>();
           Map<String, Object> headMap = new LinkedHashMap<>();
           headMap.put("ID", 1);
           headMap.put("学生成绩", 3);
           mergeHeads.add(headMap);
   
           headMap = new LinkedHashMap<>();
           headMap.put("ID", 1);
           headMap.put("姓名", 1);
           headMap.put("成绩", 2);
           mergeHeads.add(headMap);
   
           headMap = new LinkedHashMap<>();
           headMap.put("id", "ID");
           headMap.put("name", "姓名");
           headMap.put("chineseScore", "语文");
           headMap.put("mathScore", "数学");
           mergeHeads.add(headMap);
           return mergeHeads;
       }
   
       private Map<String, Map<String, String>> getMergeHeadMap() {
           Map<String, Map<String, String>> mergeHeadMap = new LinkedHashMap<>();
           Map<String, String> headMap = new LinkedHashMap<>();
           headMap.put("id", "ID");
           mergeHeadMap.put("序号", headMap);
   
           headMap = new LinkedHashMap<>();
           headMap.put("name", "姓名");
           mergeHeadMap.put("姓名", headMap);
   
           headMap = new LinkedHashMap<>();
           headMap.put("chineseScore", "语文");
           headMap.put("mathScore", "数学");
           mergeHeadMap.put("成绩", headMap);
           return mergeHeadMap;
       }
   
       private List<Student> getScores() {
           List<Student> students = new ArrayList<>();
           try {
               Student student = new Student(1, "张三三三三三三三三三三三", 99, 100);
               students.add(student);
               student = new Student(2, "李四", 99, 79);
               students.add(student);
               student = new Student(3, "王五", 80, 80);
               students.add(student);
               student = new Student(4, "赵六", 60, 59);
               students.add(student);
           } catch (Exception e) {
               //
           }
           return students;
       }
   
       private List<Student> getStudents() {
           SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
           List<Student> students = new ArrayList<>();
           try {
               Student student = new Student(1, "张三", sdf.parse("1993-01-01"));
               students.add(student);
               student = new Student(2, "李四", sdf.parse("1994-01-01"));
               students.add(student);
               student = new Student(3, "王五", sdf.parse("1995-01-01"));
               students.add(student);
               student = new Student(4, "赵六", sdf.parse("1996-01-01"));
               students.add(student);
           } catch (Exception e) {
               //
           }
           return students;
       }
   
       private List<Teacher> getTeachers() {
           List<Teacher> teachers = new ArrayList<>();
           Teacher teacher = new Teacher(1, "张三", "语文");
           teachers.add(teacher);
           teacher = new Teacher(2, "李四", "数学");
           teachers.add(teacher);
           teacher = new Teacher(3, "王五", "音乐");
           teachers.add(teacher);
           teacher = new Teacher(4, "赵六", "体育");
           teachers.add(teacher);
           return teachers;
       }

```
