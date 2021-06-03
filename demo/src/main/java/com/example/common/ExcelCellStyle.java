package com.example.common;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Data;

@Data
public class ExcelCellStyle {
    private XSSFWorkbook wb;
    // 「届出」シートの社員テーブルのセルスタイル
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
        // セルスタイルの作成
        this.notificationEmployeeInfoTable      = wb.createCellStyle();
        this.notificationTableHeaderNorm        = wb.createCellStyle();
        this.notificationTableHeaderLeftDash    = wb.createCellStyle();
        this.notificationTableHeaderRightDash   = wb.createCellStyle();
        this.styleNotificationTitleYear         = wb.createCellStyle();
        this.notificationFooterStyle            = wb.createCellStyle();
        // フォントの作成
        this.notificationEmployeeInfoTableFont      = wb.createFont();
        this.notificationTableHeaderNormFont        = wb.createFont();
        this.notificationTableHeaderLeftDashFont    = wb.createFont();
        this.notificationTableHeaderRightDashFont   = wb.createFont();
        this.styleNotificationTitleYearFont         = wb.createFont();
        this.notificationFooterFont                 = wb.createFont();
        this.fontNotificationTitleYear                = wb.createFont();
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
        notificationFooterFont.setColor(ExcelFont.footer);

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
        notificationFooterStyle.setFillForegroundColor(ExcelFont.footer);         


    }
}