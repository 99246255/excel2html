package com.codecraft.excel2html.utils;

import com.codecraft.excel2html.entity.ExcelTable;
import com.codecraft.excel2html.entity.ExcelTableTd;
import com.codecraft.excel2html.entity.ExcelTableTr;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;

/**
 * excel工具类
 * @author zoro
 *
 */
public class ExcelUtils {
	/*
	 * 用于将excel表格中列索引转成列号字母，从A对应1开始
	 */
	public static String indexToColumn(int index) {
        if(index <= 0){                
        	try{                     
        		throw new Exception("Invalid parameter");                 
        	}catch (Exception e) {                         
        		e.printStackTrace();                
        	}         
        }         
        index--;         
        String column = "";         
        do{                
        	if(column.length() > 0) {
                        index--;
            }
        	column = ((char) (index % 26 + (int) 'A')) + column;
            index = (int) ((index - index % 26) / 26);
        }while(index>0);
        return column;
	}
	
	/*
	 * 计算合并列的宽度度(单位为px)
	 */
	public static double getTdSpanWidth(Sheet sheet, int startCol, int endCol) {
		double tdwidth = 0;
		for (int i = startCol; i <= endCol; i++) {
			double tempwidth = sheet.getColumnWidthInPixels(i);//获得px像素
			tdwidth = tdwidth + tempwidth;

		}
		return tdwidth;
	}
	
	/*
	 * 计算合并列的高度
	 */
	public static int getTdSpanHeight(Sheet sheet, int startRow, int endRow) {
		int tdHeight = 0;
		for (int i = startRow; i <= endRow; i++) {
			int tempHeight = sheet.getRow(i).getHeight() / 32;
			tdHeight = tdHeight + tempHeight;
		}
		return tdHeight;
	}
	
	/*
	 * 计算合并列的高度
	 */
	public static String getTdSpanHeight(ExcelTable table, ExcelTableTr tr,
			ExcelTableTd td) {
		int rowspan = 0;
		if(!"".equals(td.getRowspan())){
			rowspan = Integer.parseInt(td.getRowspan());
		}
		//计算合并单元格的高度
		int height = Integer.parseInt(tr.getHeight().substring(0, tr.getHeight().indexOf("px")));
		if(rowspan > 1){
			int thisIndex = tr.getRowNum();
			for(int i=1; i < rowspan; i++){
				ExcelTableTr trBo = table.getTrMap().get(thisIndex + i);
				if(trBo != null){
					height+=Integer.parseInt(trBo.getHeight().substring(0, trBo.getHeight().indexOf("px")));
				}
			}
		}
		return height + "px";
	}
	
	/**
	 * 根据文件的路径创建Workbook对象
	 * 
	 * @param filePath
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws Exception
	 */
	public static Workbook getExcelWorkBook(InputStream ins) throws Exception {
		Workbook book = null;
		try {
			book = WorkbookFactory.create(ins);
		}catch(IllegalArgumentException e){
			throw new Exception("Excel打开出错!");
		}
		return book;
	}
	
	/**
	 * 判断Excel版本
	 * @param filePath
	 * @return
	 */
	public static boolean isExcelHSSF(InputStream input) {
		Workbook workbook = null;
		try {
			workbook = new HSSFWorkbook(input);
		} catch (Exception e) {
			return false;
		} finally{
			if(workbook != null){
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		return true;
	}
	
	/**
	 * 转换为标准颜色
	 * @param hc
	 * @return
	 */
	public static String convertToStardColor(HSSFColor hc) {
		StringBuffer sb = new StringBuffer("");
		if (hc != null) {
			if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
				return null;
			}
			sb.append("#");
			for (int i = 0; i < hc.getTriplet().length; i++) {
				sb.append(fillWithZero(Integer.toHexString(hc
					.getTriplet()[i])));
			}
		}
		return sb.toString();
	}
	
	/**
	 * 十六进制补0
	 * @param str
	 * @return
	 */
	public static String fillWithZero(String str) {
		if (str != null && str.length() < 2) {
			return "0" + str;
		}
		return str;
	}
	
	/**
	 * 单元格水平对齐
	 * @param alignment
	 * @return
	 */
	public static String convertAlignToHtml(short alignment) {
		String align = "";
		switch (alignment) {
			case CellStyle.ALIGN_LEFT:
				align = "left";
				break;
			case CellStyle.ALIGN_CENTER:
				align = "center";
				break;
			case CellStyle.ALIGN_RIGHT:
				align = "right";
				break;
			default:
				break;
		}
		return align;
	}
	
	/**
	 * 单元格垂直对齐
	 * @param verticalAlignment
	 * @return
	 */
	public static String convertVerticalAlignToHtml(short verticalAlignment) {
		String valign = "";
		switch (verticalAlignment) {
			case CellStyle.VERTICAL_BOTTOM:
				valign = "bottom";
				break;
			case CellStyle.VERTICAL_CENTER:
				valign = "middle";
				break;
			case CellStyle.VERTICAL_TOP:
				valign = "top";
				break;
			case CellStyle.VERTICAL_JUSTIFY:
				valign = "baseline";
				break;
			default:
				break;
		}
		return valign;
	}
	
	/**
	 * 获取border样式
	 * @param borderType
	 * @param colorType
	 * @return
	 */
	public static String getBorderStyle(short borderType, short colorType) {
		String html = "none";
		switch (borderType) {
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_NONE:
			html = "none";
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_THIN:
			html = "1px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM:
			html = "2px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_DASHED:
			html = "1px dashed " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_HAIR:
			html = "1px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_THICK:
			html = "5px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_DOUBLE:
			html = "double solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_DOTTED:
			html = "1px dotted " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASHED:
			html = "3px dashed " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT:
			html = "1px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT:
			html = "3px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_DASH_DOT_DOT:
			html = "1px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
			html = "3px solid " + getBorderColor(colorType);
			break;
		case org.apache.poi.ss.usermodel.CellStyle.BORDER_SLANTED_DASH_DOT:
			html = "1px solid " + getBorderColor(colorType);
			break;
		default:
			break;
		}
		return html;
	}
	
	/**
	 * 获取border颜色
	 * @param bordercolor
	 * @return
	 */
	public static String getBorderColor(short bordercolor) {
		String type = "black";
		if(bordercolor == HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.ROYAL_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.CORAL.getIndex()) {
			type = "coral";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.ORCHID.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.MAROON.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.LEMON_CHIFFON.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.CORNFLOWER_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.WHITE.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.LAVENDER.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.PALE_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.LIGHT_TURQUOISE.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.LIGHT_YELLOW.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.TAN.getIndex()) {
			type = "tan";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.ROSE.getIndex()) {
			type = "rose";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.PLUM.getIndex()) {
			type = "black";
		}else if(bordercolor ==HSSFColor.HSSFColorPredefined.SKY_BLUE.getIndex()) {
			type = "blue";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.TURQUOISE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.BRIGHT_GREEN.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.YELLOW.getIndex()) {
			type = "yellow";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.GOLD.getIndex()) {
			type = "gold";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.PINK.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.VIOLET.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.AQUA.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.SEA_GREEN.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.LIME.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.LIGHT_ORANGE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.RED.getIndex()) {
			type = "red";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex()) {
			type = "grey";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.BLUE.getIndex()) {
			type = "blue";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.TEAL.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.GREEN.getIndex()) {
			type = "green";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.DARK_YELLOW.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.ORANGE.getIndex()) {
			type = "orange";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.DARK_RED.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.INDIGO.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.DARK_BLUE.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.DARK_TEAL.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.DARK_GREEN.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.OLIVE_GREEN.getIndex()) {
			type = "black";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.BROWN.getIndex()) {
			type = "brown";
		}else if(bordercolor == HSSFColor.HSSFColorPredefined.BLACK.getIndex()) {
			type = "black";
		}
		return type;
	}
}
