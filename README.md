# excel2html
引用 https://github.com/wanglong1615/excel2html的

Excel2Html顾名思义将Excel转换为html.(注意:部分代码参照于别人.)
目前仅支持Excel2003格式.

如果表格的样式不对,那么请注意以下两点:
1.hssf默认字体为10px宋体.
sheet.getColumnWidthInPixels : 
Please note, that this method works correctly only for workbooks with the default font size (Arial 10pt for .xls and Calibri 11pt for .xlsx).
2.excel中的内容不要充满整个单元格.(单元格内容过大会将table撑大)


修改内容如下：
1. 升级了POI版本
2. 生成的表格居中显示
3. 参考POI examples toHTML, 重写获取内容的方法， 之前的内容数字不支持百分比，不支持公式
4. 修改FilePrint
## main test
                String excelPath = "F:\\1.xls";
                String htmlPath = "f:\\html\\2.html";
                ConvertConfig config = new ConvertConfig().setHtmlPrint(new FilePrint(htmlPath))
                        .setMaxRowNum(500).setMaxCellNum(500).setExcelType("HSSF");
                Excel2Html excel2Html = new Excel2Html(config);
                excel2Html.conver(excelPath, 0);
		
		
目前针对自己的项目够用了，有时间以后再完善修改
  
