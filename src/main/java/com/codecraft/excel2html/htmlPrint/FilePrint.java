package com.codecraft.excel2html.htmlPrint;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;

/**
 * 文件输出
 * @author zoro
 *
 */
public class FilePrint implements IHtmlPrint {
	
	private String fileName;

	public FilePrint(String fileName){
	 	this.fileName = fileName;
	}
	

	public void print(String htmlContent) throws Exception {
		File printFile = new File(fileName);
		if(!printFile.exists()){
			printFile.createNewFile();
		}
		FileOutputStream fos = null;
		OutputStreamWriter osw = null;
		BufferedWriter bw = null;
		try{
			fos = new FileOutputStream(printFile);
			osw =  new OutputStreamWriter(fos);
			bw = new BufferedWriter(osw);
			bw.write(htmlContent);
			bw.flush();
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			if(bw != null){
				bw.close();
			}
			if(osw != null){
				osw.close();
			}
			if(fos != null){
				fos.close();
			}
		}
	}
}
