**ExcelTool**中的`ExcelUtil`是`poi`封装好的excel导出工具。
- **单表导出**：将一个表格导出；对应方法名`downloadTwoTable2Excel(response, content1, title1, content2, title2, fileName)`
- **双表导出**：将二个表格导出到一个`excel`中，对应的方法名为`downloadTwoTable2Excel(response, content1, title1, content2, title2, fileName);`<br/>
  解释一下参数：<br/>
  `response` 是`HttpServletResponse`类型<br/>
  `content1` `content2`是`List<LinkedHashMap<String,Object>>`类型<br/>
  `title1` `title` 是`excel`的表头<br/>
  `fileName`是保存的文件名<br/>
**下载个2个表格的效果图：**<br/>
![image](https://github.com/12-09/ExcelTool/blob/master/ExcelTool/WebContent/WEB-INF/images/ExcelTool.png)
