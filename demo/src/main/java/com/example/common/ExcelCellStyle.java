package com.example.common;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Data;

@Data
public class ExcelCellStyle {
    private XSSFWorkbook wb;
    
    // セルスタイルテンプレート
    // 枠線、水平・垂直位置センタリング
    private XSSFCellStyle centeredThinBorderStyle;
    // 薄い枠線でスタイルを初期化




    // 「届出」シート
    // 社員テーブルのセルスタイル
    private XSSFCellStyle notificationEmployeeInfoTable;
    private XSSFCellStyle notificationTableHeaderNorm;
    private XSSFCellStyle notificationTableHeaderLeftDash;
    private XSSFCellStyle notificationTableHeaderRightDash;
    private XSSFCellStyle styleNotificationTitleYear;
    private XSSFCellStyle notificationFooterStyle;

    private XSSFFont notificationEmployeeInfoTableFont;
    private XSSFFont notificationTableHeaderNormFont;
    private XSSFFont notificationTableHeaderLeftDashFont;
    private XSSFFont notificationTableHeaderRightDashFont;
    private XSSFFont styleNotificationTitleYearFont;
    private XSSFFont notificationFooterFont;
    private XSSFFont fontNotificationTitleYear;

    /**
     * コンストラクタ
     */
    public ExcelCellStyle(XSSFWorkbook wb){
        // ワークブック
        this.wb = wb;
        // セルスタイル様式の初期化
        centeredThinBorderStyle = wb.createCellStyle();
        centeredThinBorderStyle.setAlignment(HorizontalAlignment.CENTER);
        centeredThinBorderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        centeredThinBorderStyle.setBorderBottom(BorderStyle.THIN);
        centeredThinBorderStyle.setBorderTop(BorderStyle.THIN);
        centeredThinBorderStyle.setBorderLeft(BorderStyle.THIN);
        centeredThinBorderStyle.setBorderRight(BorderStyle.THIN);
        // 「届出」
        // セルスタイルの作成
        notificationEmployeeInfoTable = wb.createCellStyle();
        notificationEmployeeInfoTable.cloneStyleFrom(centeredThinBorderStyle);
        notificationTableHeaderNorm = wb.createCellStyle();
        notificationTableHeaderNorm.cloneStyleFrom(centeredThinBorderStyle);
        notificationTableHeaderLeftDash = wb.createCellStyle();
        notificationTableHeaderLeftDash.cloneStyleFrom(centeredThinBorderStyle);
        notificationTableHeaderRightDash = wb.createCellStyle();
        notificationTableHeaderRightDash.cloneStyleFrom(centeredThinBorderStyle);
        notificationFooterStyle = wb.createCellStyle();
        notificationFooterStyle.cloneStyleFrom(centeredThinBorderStyle);
        // YYYY年度 届出タイトル部分のスタイル
        styleNotificationTitleYear = wb.createCellStyle();
        fontNotificationTitleYear.setBold(true);
        fontNotificationTitleYear.setFontHeightInPoints((short)16);;
        fontNotificationTitleYear.setFontName("ＭＳ Ｐ明朝");

        // フォントの作成
        notificationEmployeeInfoTableFont      = wb.createFont();
        notificationTableHeaderNormFont        = wb.createFont();
        notificationTableHeaderLeftDashFont    = wb.createFont();
        notificationTableHeaderRightDashFont   = wb.createFont();
        styleNotificationTitleYearFont         = wb.createFont();
        notificationFooterFont                 = wb.createFont();
        fontNotificationTitleYear              = wb.createFont();

        // フォントの設定をセット
        notificationEmployeeInfoTableFont.setFontName("ＭＳ ゴシック");;
        notificationTableHeaderNormFont.setBold(true);
        notificationTableHeaderNormFont.setFontName("ＭＳ Ｐゴシック");
        // notificationTableHeaderLeftDashFont.;
        // notificationTableHeaderRightDashFont;
        // styleNotificationTitleYearFont      ;
        notificationFooterFont.setFontName("Century");
        notificationFooterFont.setFontHeightInPoints((short)16);
        notificationFooterFont.setBold(true);
        notificationFooterFont.setColor(ExcelColor.footer);

        fontNotificationTitleYear.setBold(true);
        fontNotificationTitleYear.setFontHeightInPoints((short)16);;
        fontNotificationTitleYear.setFontName("ＭＳ Ｐ明朝");
        styleNotificationTitleYear.setFont(fontNotificationTitleYear);



        // セルスタイルにフォント設定をセット
        // notificationEmployeeInfoTable;   
        // notificationTableHeaderNorm     
        // notificationTableHeaderLeftDash 
        // notificationTableHeaderRightDash
        // styleNotificationTitleYear      
        notificationFooterStyle.setFillForegroundColor(ExcelColor.footer);         


    }
}