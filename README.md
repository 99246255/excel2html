# excel2html
将Excel转为html,部分代码来源于第三方。
 	//main test
  	public static void main(String[] args) throws Exception{
			File file = new File("C:/Temp/1.xls");
			ConvertConfig config = new ConvertConfig();
			config.setHtmlPrint(new FilePrint("test"));
			config.setMaxRowNum(500).setMaxCellNum(500).setExcelType("HSSF");
			Excel2Html excel2Html = new Excel2Html(config);
			FileInputStream fis = new FileInputStream(file);
			excel2Html.conver(fis, 0);
			fis.close();
  	} 
    
    The HTML file stored in C:/Temp.
  
