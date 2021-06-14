package com.example.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 既存のエクセルから幅や高さを取得するためのクラス
 */
public class GetW {

    public static void main(String[] args) throws Exception{
        
        String tempFilename = "勤怠報告書（2020年度_大野原信）1.xlsm";
        
        Path dir = Path.of("C:/Users/yumasky/Desktop/work/VScode/managing-atendance/apachepoi/output/");
        File file = new File(dir.toFile(), tempFilename);
        FileInputStream is = null;
        try {
            is = new FileInputStream(file);
            XSSFWorkbook wb = (XSSFWorkbook)WorkbookFactory.create(is);
            // XSSFSheet todokedeSheet = wb.getSheet("届出");
            XSSFSheet todokedeSheet = wb.getSheet("4月");

            System.out.println("********幅を取得********");
            for(int i = 0; i < 17; i++){
                // System.out.println((i+1) +"列目の幅: " + todokedeSheet.getColumnWidth(i));
                System.out.println(todokedeSheet.getColumnWidth(i));
            }
            System.out.println("********高さを取得********");
            // for(int i = 0; i < 48; i++){
            for(int i = 0; i < 53; i++){
                // System.out.println((i+1) +"行目の高さ: " + todokedeSheet.getRow(i).getHeight());
                System.out.println(todokedeSheet.getRow(i).getHeight());
            }

        } catch (IOException e) {
            System.out.println(e.toString());
        } catch (Exception e) {
            System.out.println(e.toString());
        } finally {
            try{
                is.close();
            } catch (IOException e){
                System.out.println(e.toString());
            }
        }
    }
}
