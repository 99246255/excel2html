package com.codecraft.excel2html;

import org.kamranzafar.jtar.TarEntry;
import org.kamranzafar.jtar.TarOutputStream;

import java.io.*;

/**
 * 打包测试
 */
public class TarTest {
    public static void main(String[] args) throws Exception {
        FileOutputStream dest = new FileOutputStream( "f:/1.zip" );
        TarOutputStream out = new TarOutputStream( new BufferedOutputStream( dest ) );
        // Files to tar
        File[] filesToTar=new File[2];
        filesToTar[0]=new File("f:/html/1.html");
        filesToTar[1]=new File("f:/html/2.html");
        for(File f:filesToTar){
            out.putNextEntry(new TarEntry(f, f.getName()));
            BufferedInputStream origin = new BufferedInputStream(new FileInputStream( f ));
            int count;
            byte data[] = new byte[2048];
            while((count = origin.read(data)) != -1) {
                out.write(data, 0, count);
            }
            out.flush();
            origin.close();
        }
        out.close();
    }
}
