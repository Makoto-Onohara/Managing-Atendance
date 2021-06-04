package com.example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import com.example.common.ExcelCellStyle;
import com.example.common.ExcelColor;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException
    {
        

        // 社員情報
        String name = "大野原信";     // 名前
        String employeeCode = "00249"; // 社員番号

        String[] notificationHeaderStrings = {
            "提出", "", "区分", "開始", "", "", "終了", "", "","日数","承認","備考(自由等)",
            "月","日", "", "月","日","時間","月","日","時間", "", "", ""
        };
        
        // エクセルを保存するディレクトリ
        Path dir = Path.of("C:/Users/yumasky/Desktop/work/VScode/managing-atendance/apachepoi/");
        // define a file name
        File file = new File(dir.toFile(), "勤怠報告書（2021年度_" + name +"）.xlsx");
        
        FileOutputStream fos = null;

        int rowSizeNotification = 48;   // 届出シートの行数
        int colSizeNotification = 12;    // 届出シートの列数
        int rowSize = 41;               // ○月シートの行数
        int colSize = 9;                // ○月シートの列数
        
        
        
        // // create a new file
        // FileOutputStream out = new FileOutputStream("workbook.xlsx");
        
        // create a new workbook
        XSSFWorkbook wb = new XSSFWorkbook();
        
        // セルスタイル保持クラスを生成
        // ExcelCellStyle cellStyle = new ExcelCellStyle(wb);
        // create a new sheet
        List<XSSFSheet> sheetList = new ArrayList<>();
        sheetList.add(wb.createSheet("届出"));
        sheetList.add(wb.createSheet("5月"));
        XSSFSheet sheetNotification = sheetList.get(0);
        XSSFSheet sheet = sheetList.get(1);
        // create a row
        XSSFRow row = null;
        XSSFCell cell = null;
        // シート［届出」のセルスタイル
        // 社員情報のスタイル
        XSSFCellStyle notificationEmployeeInfoTable = wb.createCellStyle();
        XSSFFont notificationEmployeeInfoTableFont = wb.createFont();
        notificationEmployeeInfoTable.setAlignment(HorizontalAlignment.CENTER);
        notificationEmployeeInfoTable.setVerticalAlignment(VerticalAlignment.CENTER);
        notificationEmployeeInfoTable.setBorderBottom(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderTop(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderLeft(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderRight(BorderStyle.THIN);
        notificationEmployeeInfoTable.setFont(notificationEmployeeInfoTableFont);
        // ヘッダー部分のノーマルスタイル
        XSSFCellStyle notificationTableHeaderNorm = wb.createCellStyle();
        XSSFFont notificationTableHeaderNormFont = wb.createFont();
        notificationTableHeaderNorm.cloneStyleFrom(notificationEmployeeInfoTable);
        notificationTableHeaderNorm.setFillForegroundColor(ExcelColor.blueNote);
        notificationTableHeaderNorm.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        notificationTableHeaderNormFont.setBold(true);
        notificationTableHeaderNormFont.setFontName("ＭＳ Ｐゴシック");
        notificationTableHeaderNorm.setFont(notificationTableHeaderNormFont);
        // ヘッダー部分の左破線
        XSSFCellStyle notificationTableHeaderLeftDash = wb.createCellStyle();
        notificationTableHeaderLeftDash.cloneStyleFrom(notificationTableHeaderNorm);
        // ヘッダー部分の右破線
        XSSFCellStyle notificationTableHeaderRightDash = wb.createCellStyle();
        notificationTableHeaderRightDash.cloneStyleFrom(notificationTableHeaderNorm);
        // セルの結合
        sheetNotification.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(2, 2, 2, 3));
        sheetNotification.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(3, 3, 2, 3));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 3, 5));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 6, 8));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 9, 9));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 10, 10));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 11, 11));


    // ココから「届出」シート
        for( int i = 0; i < rowSizeNotification; i++ ){
            row = sheetNotification.createRow(i);
            for(int j = 0; j < colSizeNotification; j++){
                row.createCell(j);
            }
        }
        // YYYY年度 届出
        row = sheetNotification.getRow(0);
        cell = row.getCell(0);
        cell.setCellValue("2021年度 届出");
        XSSFCellStyle styleNotificationTitleYear = wb.createCellStyle();
        XSSFFont fontNotificationTitleYear = wb.createFont();
        fontNotificationTitleYear.setBold(true);
        fontNotificationTitleYear.setFontHeightInPoints((short)16);;
        fontNotificationTitleYear.setFontName("ＭＳ Ｐ明朝");
        styleNotificationTitleYear.setFont(fontNotificationTitleYear);
        cell.setCellStyle(styleNotificationTitleYear);
        cell.setCellStyle(styleNotificationTitleYear);
        // 社員情報
        row = sheetNotification.getRow(2);
        cell = row.getCell(0);
        cell.setCellValue("社員番号");
        cell = row.getCell(2);
        cell.setCellValue("氏名");
        row = sheetNotification.getRow(3);
        cell = row.getCell(0);
        cell.setCellValue(employeeCode);
        cell = row.getCell(2);
        cell.setCellValue(name);
        for(int i = 2; i < 4; i++){
            for(int j = 0; j < 4; j++){
                sheetNotification.getRow(i).getCell(j).setCellStyle(notificationEmployeeInfoTable);
            }
        }
        // テーブル部分作成
        for(int i = 5; i < 47; i++){
            row = sheetNotification.getRow(i);
            for(int j = 0; j < 12; j++){
                cell = row.getCell(j);
                if(i == 5){
                    cell.setCellValue(notificationHeaderStrings[j]);
                    cell.setCellStyle(notificationTableHeaderNorm);
                } else if(i == 6){
                    cell.setCellValue(notificationHeaderStrings[j + 12]);    
                    cell.setCellStyle(notificationTableHeaderNorm);
                } else {
                    cell.setCellStyle(notificationEmployeeInfoTable);
                }


            }

        }
        // フッター部分
        row = sheetNotification.getRow(47);
        cell = row.getCell(0);
        cell.setCellValue("MicroMagic INC.");
        XSSFCellStyle notificationFooterStyle = wb.createCellStyle();
        XSSFFont notificationFooterFont = wb.createFont();
        notificationFooterFont.setFontName("Century");
        notificationFooterFont.setFontHeightInPoints((short)16);
        notificationFooterFont.setBold(true);
        notificationFooterFont.setColor(ExcelColor.footer);
        notificationFooterStyle.setFont(notificationFooterFont);
        notificationFooterStyle.setAlignment(HorizontalAlignment.CENTER);
        notificationFooterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // notificationFooterStyle.setFillForegroundColor(new XSSFColor(footerRGB, new DefaultIndexedColorMap()));
        cell.setCellStyle(notificationFooterStyle);
        sheetNotification.addMergedRegion(new CellRangeAddress(47,47,0,11));




    // ココから「５月」シート
        for( int i = 0; i < rowSize; i++ ){
            row = sheet.createRow(i);
            for(int j = 0; j < colSize; j++){
                row.createCell(j);
            }
        }
        // XSSFCell cell = row.createCell(0);

        // 結合する相手のセルは作成していなくても例外は発生しない
        // XSSFCell cell1 = row.createCell(1);
        // XSSFCell cell2 = row.createCell(2);
        
        // 「YYYY年M月分」
        sheet.getRow(0).getCell(0).setCellValue("2021年5月分");
        // merge cells of 年月
        // 引数の番号は0基底
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,2));

        // フォントの設定は一括で行い、セルに対してセットする
        // セルごとにフォントをセットしないといけない？
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontName("ＭＳ Ｐ明朝");
        // font.setUnderline(Font.U_SINGLE);
        font.setFontHeightInPoints((short)16);

        // CellStyle はセルごとにインスタンスを生成する必要がある
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        style.setVerticalAlignment(VerticalAlignment.CENTER);;
        row = sheet.getRow(0);
        cell = row.getCell(0);
        cell.setCellStyle(style);

        // 「勤怠報告書」
        CellStyle style2 = wb.createCellStyle(); 
        style2.setVerticalAlignment(VerticalAlignment.CENTER);;
        cell = row.createCell(3);
        cell.setCellValue("勤怠報告書");
        Font font2 = wb.createFont();
        font2.setBold(true);
        font2.setFontHeightInPoints((short)20);
        font2.setFontName("ＭＳ Ｐ明朝");
        style2.setFont(font2);
        cell.setCellStyle(style2);


        // 会社名の設定
        cell = row.createCell(7);
        cell.setCellValue("株式会社マイクロマジック");
        CellStyle style3 = wb.createCellStyle();
        Font font3 = wb.createFont();
        font3.setFontHeightInPoints((short)10);
        font3.setFontName("ＭＳ Ｐ明朝");
        style3.setFont(font3);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);;
        cell.setCellStyle(style3);

        // 「開始」
        row = sheet.getRow(1);
        cell = row.createCell(1);
        cell.setCellValue("開始");
        CellStyle style4 = wb.createCellStyle();
        Font font4 = wb.createFont();
        font4.setFontHeightInPoints((short)9);
        font4.setFontName("ＭＳ Ｐ明朝");
        style4.setFont(font4);
        style4.setVerticalAlignment(VerticalAlignment.CENTER);
        style4.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(style4);
        // 「YYYY年M月D日」
        cell = row.getCell(2);
        cell.setCellValue("2021年5月1日");
        CellStyle style4_2 = wb.createCellStyle();
        Font font4_2 = wb.createFont();
        font4_2.setBold(true);
        font4_2.setUnderline(Font.U_SINGLE);
        font4_2.setFontName("ＭＳ Ｐ明朝");
        font4_2.setFontHeightInPoints((short)9);
        style4_2.setFont(font4_2);
        style4_2.setVerticalAlignment(VerticalAlignment.CENTER);
        cell.setCellStyle(style4_2);
        


        // 「締日」
        row = sheet.getRow(2);
        cell = row.createCell(1);
        cell.setCellValue("締日");
        cell.setCellStyle(style4);
        // 「YYYY年M月D日」
        cell = row.getCell(2);
        cell.setCellValue("2021年5月31日");
        cell.setCellStyle(style4_2);


        // 「社員番号」
        row = sheet.getRow(3);
        cell = row.createCell(1);
        cell.setCellValue("社員番号");
        CellStyle style5 = wb.createCellStyle();
        Font font5 = wb.createFont();
        font5.setFontHeightInPoints((short)9);
        font5.setFontName("ＭＳ Ｐ明朝");
        style5.setFont(font5);
        style5.setVerticalAlignment(VerticalAlignment.CENTER);
        style5.setAlignment(HorizontalAlignment.CENTER);
        style5.setBorderBottom(BorderStyle.THIN);
        style5.setBorderTop(BorderStyle.THIN);
        style5.setBorderLeft(BorderStyle.THIN);
        style5.setBorderRight(BorderStyle.THIN);
        cell.setCellStyle(style5);
        // 「氏名」
        cell = row.getCell(2);
        cell.setCellValue("氏名");
        cell.setCellStyle(style5);
        // 「担当」
        cell = row.getCell(7);
        cell.setCellValue("担当");
        cell.setCellStyle(style5);
        // 「確認」
        cell = row.getCell(8);
        cell.setCellValue("確認");
        cell.setCellStyle(style5);


        // 社員番号の値
        row = sheet.getRow(4);
        cell = row.createCell(1);
        cell.setCellValue("00249");
        // セルスタイルとフォントはstyle5,font5と共通
        cell.setCellStyle(style5);
        // 氏名の値
        cell = row.getCell(2);
        cell.setCellValue("大野原　信");
        cell.setCellStyle(style5);
        // 担当欄
        cell = row.getCell(7);
        cell.setCellStyle(style5);
        // 確認欄
        cell = row.getCell(8);
        cell.setCellStyle(style5);

        
        // 「提出日」
        row = sheet.getRow(5);
        cell = row.createCell(6);
        cell.setCellValue("提出日 2021年5月31日");
        CellStyle style6 = wb.createCellStyle();
        style6.setAlignment(HorizontalAlignment.RIGHT);
        cell.setCellStyle(style6);
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 6, 8));
        
        
        // テーブルヘッダー
        row = sheet.getRow(6);
        String[] tableHeaderStrings = {
            "日付",
            "作業項目",
            "備考",
            "開始時間",
            "終了時間",
            "全時間",
            "作業時間",
            "残業時間"
        };
        
        // 「日付」
        cell = row.getCell(0);
        cell.setCellValue(tableHeaderStrings[0]);
        sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 1));
        Font fontTableHeader = wb.createFont();
        fontTableHeader.setBold(true);
        fontTableHeader.setFontHeightInPoints((short)9);
        fontTableHeader.setFontName("ＭＳ Ｐゴシック");
        XSSFCellStyle styleTableHeader = wb.createCellStyle();
        styleTableHeader.setAlignment(HorizontalAlignment.CENTER);        
        styleTableHeader.setVerticalAlignment(VerticalAlignment.CENTER);
        styleTableHeader.setFont(fontTableHeader);
        styleTableHeader.setFillForegroundColor(ExcelColor.blue);
        styleTableHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(styleTableHeader);
        cell = row.getCell(1);
        cell.setCellStyle(styleTableHeader);
        // 「作業項目」「備考」「開始時間」「終了時間」「全時間」「作業時間」「残業時間」
        for(int i = 2; i < 9; i++){
            cell = row.getCell(i);
            cell.setCellValue(tableHeaderStrings[i - 1]);
            cell.setCellStyle(styleTableHeader);
        }

        for(int i = 2; i < 9; i++){
            cell = row.getCell(i);
            cell.setCellStyle(styleTableHeader);
        }



        // カレンダー作成
        // List<XSSFCell> cellList = new ArrayList<>();
        // XSSFCell[] cellB = null;
        XSSFCell cellA = null;
        XSSFCell cellB = null;

        // LocalDate startDate;
        LocalDate startDate = LocalDate.parse("2021-05-01");
        // DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.JAPANESE);
        LocalDate date = null;
        // 平日のスタイル
        XSSFCellStyle styleWeekDay = wb.createCellStyle();
        styleWeekDay.setAlignment(HorizontalAlignment.CENTER);
        styleWeekDay.setFillForegroundColor(ExcelColor.blue);
        styleWeekDay.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontWeekDay = wb.createFont();
        fontWeekDay.setFontName("ＭＳ Ｐゴシック");
        fontWeekDay.setFontHeightInPoints((short)12);
        styleWeekDay.setFont(fontWeekDay);
        // 平日以外のスタイル
        XSSFCellStyle styleWeekEnd = wb.createCellStyle();
        styleWeekEnd.setAlignment(HorizontalAlignment.CENTER);
        styleWeekEnd.setFillForegroundColor(ExcelColor.blue);
        styleWeekEnd.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontWeekEnd = wb.createFont();
        fontWeekEnd.setColor(IndexedColors.RED.index);
        fontWeekEnd.setFontName("ＭＳ Ｐゴシック");
        fontWeekEnd.setFontHeightInPoints((short)12);
        styleWeekEnd.setFont(fontWeekEnd);
        // 右３列のスタイル
        XSSFCellStyle styleThreeRightColumn = wb.createCellStyle();
        styleThreeRightColumn.setAlignment(HorizontalAlignment.CENTER);
        styleThreeRightColumn.setFillForegroundColor(ExcelColor.yellow);
        styleThreeRightColumn.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleThreeRightColumn.setFont(fontWeekDay); // フォントは平日と同じ 

        for(int i = 0; i < 31; i++){
            row = sheet.getRow(i + 7);
            // for(int j = 0; j < 9; j++){
            //     // cellA = row.getCell(0);
            //     // cellB = row.getCell(1);
            //     cellList.add(row.getCell(j));
            
            // }
            cellA = row.getCell(0);
            cellB = row.getCell(1);
            
            // 右３列にスタイルを適用
            row.getCell(6).setCellStyle(styleThreeRightColumn);
            row.getCell(7).setCellStyle(styleThreeRightColumn);
            row.getCell(8).setCellStyle(styleThreeRightColumn);

            
            


            // 日付を埋め込む
            // 日付のフォーマットをdとする
            // DateTimeFormatter dtf = DateTimeFormatter.ofPattern("d");
            // date = LocalDate.parse(startDate.toString(), dtf);
            
            // 日付を加算する
            date = startDate.plusDays(i);
            cellA.setCellValue(date.getDayOfMonth());
            // 日付を
            cellB.setCellValue(date.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.JAPANESE));

            // セルスタイルのセット
            if(date.getDayOfWeek() == DayOfWeek.SATURDAY || date.getDayOfWeek() == DayOfWeek.SUNDAY){
                // 休日のスタイル
                cellA.setCellStyle(styleWeekEnd);
                cellB.setCellStyle(styleWeekEnd);
            } else {
                // 平日のスタイル
                cellA.setCellStyle(styleWeekDay);
                cellB.setCellStyle(styleWeekDay);
            }
        }

        // 集計行
        row = sheet.getRow(38);
        XSSFCell totalCell = row.getCell(0);
        totalCell.setCellValue("合計");
        XSSFCellStyle styleGoukei = wb.createCellStyle();
        styleGoukei.cloneStyleFrom(styleTableHeader);
        styleGoukei.setBorderTop(BorderStyle.DOUBLE);
        styleGoukei.setBorderBottom(BorderStyle.THICK);
        sheet.addMergedRegion(new CellRangeAddress(38,38,0,1));
        totalCell.setCellStyle(styleGoukei);
        row.getCell(1).setCellStyle(styleGoukei);

        XSSFCellStyle styleTotalNotEdge = wb.createCellStyle();
        styleTotalNotEdge.cloneStyleFrom(styleGoukei);
        // styleTotalNotEdge.setBorderTop(BorderStyle.DOUBLE);
        // styleTotalNotEdge.setBorderBottom(BorderStyle.THICK);
        
        for(int i = 2; i < 9; i++){
            row.getCell(i).setCellStyle(styleTotalNotEdge);
        }


        // row = sheet.getRow(7);
        // cell = row.getCell(0);
        // cell.setCellValue("test");
        // CellStyle testStyle = wb.createCellStyle();
        // // testStyle.setFillBackgroundColor((short)1000);
        // testStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // cell.setCellStyle(testStyle);

        
        // save a file
        
        try {
            // create a output stream from Path and File
            fos = new FileOutputStream(file);
            wb.write(fos);
        } catch (Exception e){
            e.printStackTrace();
        } finally {
            if( wb != null){
                wb.close();
            }
            if( fos != null) {
                fos.close();
            }
        }

        // try (FileOutputStream out = new FileOutputStream(filename)){
        //     wb.write(out);
        //     wb.close();
        // } catch (Exception e ){
        //     e.printStackTrace();
        // }




        
    }
}
