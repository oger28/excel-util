# excel-util
Excel导出工具类
 * 设计思想：
 * 1. 样式最简化原则
 * 2. 约定大于规定原则
 * <p>
 
 * 优点：
 * 1. 可任意创建sheet
 * 2. 可在任意位置创建table
 * 3. 可按顺序导出实体类任意字段
 * 4. 亦可一次性导出单sheet单表模式的Excel
 
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
        headMap.put("borthDate", "生日");
        ExcelExportUtil.exportExcel(fileName, headMap, getStudents(), response);
    }

    /**
     *  在一个 sheet 中有多张表
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
        ExcelExportUtil.exportTable(tableName, headMap, getStudents(), sheet1, workbook);

        sheetName = "教师列表";
        tableName = "教师列表";
        headMap = new LinkedHashMap<>();
        headMap.put("name", "姓名");
        headMap.put("subject", "科目");
        Sheet sheet2 = workbook.createSheet(sheetName);
        ExcelExportUtil.exportTable(tableName, headMap, getTeachers(), sheet2, workbook);

        ExcelExportUtil.exportExcel(fileName, workbook, response);
    }

    /**
     *  多sheet 导出
     */
    @GetMapping("/exportStudentsAndTeachers2")
    @ApiOperation(value = "导出学生和老师2")
    @ResponseBody
    public void exportStudentsAndTeachers2(HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        String fileName = "学生和老师列表";
        String sheetName = "学生和老师列表";
        Map<String, String> headMap = new LinkedHashMap<>();
        headMap.put("name", "姓名");
        headMap.put("birthday", "生日");
        String tableName = "学生列表";
        Sheet sheet = ExcelExportUtil.createSheet(fileName, sheetName, workbook);
        int line = ExcelExportUtil.exportTable(tableName, headMap, getStudents(), sheet, workbook);

        tableName = "教师列表";
        headMap = new LinkedHashMap<>();
        headMap.put("name", "姓名");
        headMap.put("subject", "科目");
        ExcelExportUtil.exportTable(line, tableName, headMap, getTeachers(), sheet, workbook);

        ExcelExportUtil.exportExcel(fileName, workbook, response);
    }

```
