package com.codecraft.excel2html.facade;

import com.codecraft.excel2html.config.ConvertConfig;
import com.codecraft.excel2html.cons.ExcelConstant;
import com.codecraft.excel2html.entity.ExcelTable;
import com.codecraft.excel2html.generator.HtmlGenerator;
import com.codecraft.excel2html.htmlPrint.FilePrint;
import com.codecraft.excel2html.htmlPrint.IHtmlPrint;
import com.codecraft.excel2html.parse.IExcelParse;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 用户只需调用conver方法即可
 * @author zoro
 *
 */
public class Excel2Html {
    //转换配置
    private ConvertConfig config = null;

    public Excel2Html() {
        this.config = new ConvertConfig();
    }

    public Excel2Html(ConvertConfig config) {
        this.config = config;
    }

    //main test
    public static void main(String[] args) throws Exception {
        String excelPath = "f:\\1.xls";
        String htmlPath = "f:\\html\\2.html";
        ConvertConfig config = new ConvertConfig().setHtmlPrint(new FilePrint(htmlPath))
                .setMaxRowNum(500).setMaxCellNum(500).setExcelType("HSSF");
        Excel2Html excel2Html = new Excel2Html(config);
        excel2Html.conver(excelPath, 0);
    }

    /**
     * 转换所有的sheet
     * @param is
     * @throws Exception
     */
    public void conver(InputStream is) throws Exception {
        IExcelParse excelParse = config.getExcelParse();
        //解析excel(用户可自定义)
        ExcelTable table = excelParse.parse(is, config);
        table.setConfig(config);

        //生成html页面
        HtmlGenerator htmlGener = new HtmlGenerator();
        String html = htmlGener.toHtml(table);

        //输出html(用户可自定义)
        IHtmlPrint print = config.getHtmlPrint();
        print.print(html);
    }

    /**
     * 转换指定sheet
     * @param fileName
     * @param sheetIndex
     * @throws Exception
     */
    public void conver(String fileName, int sheetIndex) throws Exception {
        FileInputStream fis = null;
        try {
            File file = new File(fileName);
            if (!file.exists()) {
                throw new Exception("file not found");
            }
            fis = new FileInputStream(file);
            if (!config.getExcelType().equalsIgnoreCase(ExcelConstant.EXCEL_TYPE_4_HSSF)) {
                throw new Exception("目前不支持'" + config.getExcelType() + "'类型!");
            }
            IExcelParse excelParse = config.getExcelParse();
            //解析excel(用户可自定义)
            ExcelTable excelTable = excelParse.parse(fis, sheetIndex, config);
            //生成html页面
            HtmlGenerator htmlGener = new HtmlGenerator();
            String html = htmlGener.toHtml(excelTable);
            //输出html(用户可自定义)
            IHtmlPrint print = config.getHtmlPrint();
            print.print(html);
        } catch (IOException e) {
            throw e;
        } finally {
            if (fis != null) {
                fis.close();
            }
        }

    }

    public ConvertConfig getConfig() {
        return config;
    }

    public void setConfig(ConvertConfig config) {
        this.config = config;
    }
}
