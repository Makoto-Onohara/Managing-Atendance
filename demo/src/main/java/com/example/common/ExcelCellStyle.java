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
    
    // フォント名定義
    final String CENTURY = "century"; // 数字用フォント
    final String PMINCHO = "ＭＳ Ｐ明朝"; // 文字用フォント


    /**
     * CellStyleテンプレート
     * 枠線： 四方
     * センタリング： 水平・垂直
     */
    public XSSFCellStyle centeredThinBorderStyle;

    //////////////////////////
    // 「届出」シートのスタイル
    //////////////////////////

    /**
     * 社員テーブル
     */
    public XSSFCellStyle notificationEmployeeInfoTable;

    /**
     * テーブルヘッダーノーマル
     */
    public XSSFCellStyle notificationTableHeaderNorm;

    /**
     * 届出シートのテーブルヘッダー
     * 左破線
     */
    public XSSFCellStyle notificationTableHeaderLeftDash;

    /**
     * 届出シートのテーブルヘッダー
     * 右破線
     */
    public XSSFCellStyle notificationTableHeaderRightDash;

    /**
     * 届出シートの年度タイトル
     */
    public XSSFCellStyle notificationTitleYear;

    /**
     * 届出シートのフッタースタイル
     */
    public XSSFCellStyle notificationFooterStyle;

    //////////////////////
    // 「届出」シートのフォント
    //////////////////////

    /**
     * Fontのテンプレート
     * フォント名：MS P明朝
     */
    public XSSFFont centeredThinBorderFont;

    /**
     * 届出シートの社員情報フォント
     */
    public XSSFFont notificationEmployeeInfoTableFont;     // 社員情報のフォント

    /**
     * 届出シートのテーブルノーマルフォント
     */
    public XSSFFont notificationTableHeaderNormFont;       // テーブルノーマルフォント

    /**
     * 届出シートのテーブル左破線フォント
     */
    public XSSFFont notificationTableHeaderLeftDashFont;   // テーブル左破線フォント

    /**
     * 届出シートのテーブル右破線フォント
     */
    public XSSFFont notificationTableHeaderRightDashFont;  // テーブル右破線フォント

    /**
     * 届出シートの年度タイトルフォント
     */
    public XSSFFont notificationTitleYearFont;             // 年度タイトルフォント

    /**
     * 届出シートのフッターフォント
     */
    public XSSFFont notificationFooterFont;                // フッターフォント

    /**
     * コンストラクタ
     */
    public ExcelCellStyle(XSSFWorkbook wb){
        // ワークブック
        this.wb = wb;
        // 初期化
        this.init();
    }


    /**
     * 初期化処理
     */
    private void init(){
                // テンプレートスタイル作成
                centeredThinBorderStyle = wb.createCellStyle();
                centeredThinBorderStyle.setAlignment(HorizontalAlignment.CENTER);
                centeredThinBorderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                centeredThinBorderStyle.setBorderTop(BorderStyle.THIN);
                centeredThinBorderStyle.setBorderLeft(BorderStyle.THIN);
                centeredThinBorderStyle.setBorderRight(BorderStyle.THIN);
                centeredThinBorderStyle.setBorderBottom(BorderStyle.THIN);
                // テンプレートフォント作成
                centeredThinBorderFont.setFontName(PMINCHO);
                centeredThinBorderStyle.setFont(centeredThinBorderFont);


                // 「届出」
                // セルスタイルの作成
                // notificationEmployeeInfoTable = wb.createCellStyle();
                notificationEmployeeInfoTable.cloneStyleFrom(centeredThinBorderStyle);
                // notificationTableHeaderNorm = wb.createCellStyle();
                notificationTableHeaderNorm.cloneStyleFrom(centeredThinBorderStyle);
                // notificationTableHeaderLeftDash = wb.createCellStyle();
                notificationTableHeaderLeftDash.cloneStyleFrom(centeredThinBorderStyle);
                // notificationTableHeaderRightDash = wb.createCellStyle();
                notificationTableHeaderRightDash.cloneStyleFrom(centeredThinBorderStyle);
                // notificationFooterStyle = wb.createCellStyle();
                notificationFooterStyle.cloneStyleFrom(centeredThinBorderStyle);
                // YYYY年度 届出タイトル部分のスタイル
                notificationTitleYear = wb.createCellStyle();
                // notificationTitleYearFont.setFontName(PMINCHO);
        

                /**
                 * 届出シートのフォント作成
                 */
                notificationEmployeeInfoTableFont      = new XSSFFont();
                notificationTableHeaderNormFont        = new XSSFFont();
                notificationTableHeaderLeftDashFont    = new XSSFFont();
                notificationTableHeaderRightDashFont   = new XSSFFont();
                notificationTitleYearFont              = new XSSFFont();
                notificationFooterFont                 = new XSSFFont();
                // 社員テーブル
                notificationEmployeeInfoTableFont.setFontName(PMINCHO);;
                notificationEmployeeInfoTable.setFont(notificationEmployeeInfoTableFont);
                // ヘッダーノーマル
                notificationTableHeaderNormFont.setBold(true);
                notificationTableHeaderNormFont.setFontName(PMINCHO);
                notificationTableHeaderNorm.setFont(notificationTableHeaderNormFont);
                // 年度
                notificationTitleYearFont.setBold(true);
                notificationTitleYearFont.setFontName(PMINCHO);
                notificationTitleYearFont.setFontHeightInPoints((short)16);;
                notificationTitleYear.setFont(notificationTitleYearFont);
                // フッター
                notificationFooterFont.setBold(true);
                notificationFooterFont.setFontName(PMINCHO);
                notificationFooterFont.setFontHeightInPoints((short)16);
                notificationFooterFont.setColor(ExcelColor.TODOKEDE_FOOTER);
                notificationFooterStyle.setFont(notificationFooterFont);
        
    }
}