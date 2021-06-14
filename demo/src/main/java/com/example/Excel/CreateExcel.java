package com.example.Excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.sql.Date;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import com.example.common.ExcelColor;
import com.example.common.FurikaeHeight;
import com.example.common.FurikaeWidth;
import com.example.common.TodokedeHeight;
import com.example.common.TodokedeWidth;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTextBox;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {
    // 何年度のエクセル
    Integer year = 2021;

    // フォント名定義
    String CENTURY = "century"; // 数字用フォント
    String PMINCHO = "ＭＳ Ｐ明朝"; // 文字用フォント

    // 社員情報
    String name = "大野原　信";
    String nameforsheet = "大野原信"; // 名前
    String employeeCode = "00249"; // 社員番号

    String[] notificationHeaderStrings = { "提出", "", "区分", "開始", "", "", "終了", "", "", "日数", "承認", "備考(事由等)", "月", "日",
            "", "月", "日", "時間", "月", "日", "時間", "", "", "" };
    String[] articleNoteStrings = { "有給休暇", "振替休日", "生理休暇", "慶弔休暇", "特別休暇", "欠勤", "遅刻", "早退", "私用外出", "出張", "振替出勤",
            "住所変更", "結婚", "出生", "その他" };
    String[] articleNoteStatStrings = { "未", "済", "事後" };
    // エクセルを保存するディレクトリ
    Path dir = Path.of("C:/Users/yumasky/Desktop/work/VScode/managing-atendance/apachepoi/output/");
    // define a file name
    File file = new File(dir.toFile(), "勤怠報告書（" + year +"年度_" + nameforsheet + "）.xlsx");

    FileOutputStream fos = null;

    int rowSizeNotification = 48; // 届出シートの行数
    int colSizeNotification = 12; // 届出シートの列数
    int rowSize = 41;               // ○月シートの行数
    int colSize = 9;                // ○月シートの列数

    XSSFWorkbook wb;
    List<XSSFSheet> sheetList;
    XSSFRow row = null;
    XSSFCell cell = null;

    XSSFCellStyle notificationEmployeeInfoTable;
    XSSFCellStyle notificationTableHeaderNorm;
    XSSFCellStyle notificationTableHeaderLeftDash;
    XSSFCellStyle notificationTableHeaderRightDash;
    XSSFCellStyle notificationTableHeaderBothDash;
    XSSFCellStyle notificationTableNorm;
    XSSFCellStyle notificationTableLeftDash;
    XSSFCellStyle notificationTableRightDash;
    XSSFCellStyle notificationTableBothDash;
    
    /**
     * エクセルの作成
     * @throws IOException
     */
    public void createExcel() throws IOException {

        // ワークブックの作成
        this.createWorkbook(); 
        // 届出シートの作成
        this.createTodokede();
        // 振替出勤管理表シートの作成
        this.createFurikae();
        // 月ごとの報告書シート作成
        this.createMonthly();

        try {
            // create a output stream from Path and File
            fos = new FileOutputStream(file);
            wb.write(fos);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (wb != null) {
                wb.close();
            }
            if (fos != null) {
                fos.close();
            }
        }


    }
    
    /**
     * ワークブックとシート作成する
     */
    void createWorkbook() {
        wb = new XSSFWorkbook();
        sheetList = new ArrayList<>();
        // 届出シート作成
        sheetList.add(wb.createSheet("届出"));
        // 振替出勤管理表シート作成
        sheetList.add(wb.createSheet("振替出勤管理表"));
        // // 報告書シート作成
        // sheetList.add(wb.createSheet("5月"));
    }


    /**
     * 届出シート作成する
     */

    private void createTodokede() {
        
        // 届出シートを取得
        XSSFSheet sheetNotification = sheetList.get(0);
        

        // シート［届出」のセルスタイル
        // 社員情報のスタイル
        notificationEmployeeInfoTable = wb.createCellStyle();
        notificationEmployeeInfoTable.setAlignment(HorizontalAlignment.CENTER);
        notificationEmployeeInfoTable.setVerticalAlignment(VerticalAlignment.CENTER);
        notificationEmployeeInfoTable.setBorderBottom(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderTop(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderLeft(BorderStyle.THIN);
        notificationEmployeeInfoTable.setBorderRight(BorderStyle.THIN);
        XSSFFont notificationEmployeeInfoTableFont = wb.createFont();
        notificationEmployeeInfoTableFont.setFontName(PMINCHO);
        notificationEmployeeInfoTable.setFont(notificationEmployeeInfoTableFont);
        // ヘッダー部分のノーマルスタイル
        notificationTableHeaderNorm = wb.createCellStyle();
        XSSFFont notificationTableHeaderNormFont = wb.createFont();
        notificationTableHeaderNorm.cloneStyleFrom(notificationEmployeeInfoTable);
        notificationTableHeaderNorm.setFillForegroundColor(ExcelColor.BLUE_NOTE);
        notificationTableHeaderNorm.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        notificationTableHeaderNormFont.setBold(true);
        notificationTableHeaderNormFont.setFontName(PMINCHO);
        notificationTableHeaderNorm.setFont(notificationTableHeaderNormFont);
        // ヘッダー部分の左破線
        notificationTableHeaderLeftDash = wb.createCellStyle();
        notificationTableHeaderLeftDash.cloneStyleFrom(notificationTableHeaderNorm);
        notificationTableHeaderLeftDash.setBorderLeft(BorderStyle.DASHED);
        // ヘッダー部分の右破線
        notificationTableHeaderRightDash = wb.createCellStyle();
        notificationTableHeaderRightDash.cloneStyleFrom(notificationTableHeaderNorm);
        notificationTableHeaderRightDash.setBorderRight(BorderStyle.DASHED);
        // ヘッダー部分の両側破線
        notificationTableHeaderBothDash = wb.createCellStyle();
        notificationTableHeaderBothDash.cloneStyleFrom(notificationTableHeaderNorm);
        notificationTableHeaderBothDash.setBorderLeft(BorderStyle.DASHED);
        notificationTableHeaderBothDash.setBorderRight(BorderStyle.DASHED);
        notificationTableNorm = wb.createCellStyle();
        notificationTableNorm.cloneStyleFrom(notificationEmployeeInfoTable);
        // ヘッダー以外の左破線
        notificationTableLeftDash = wb.createCellStyle();
        notificationTableLeftDash.cloneStyleFrom(notificationTableHeaderLeftDash);
        notificationTableLeftDash.setFillPattern(FillPatternType.NO_FILL);
        // ヘッダー以外の右破線
        notificationTableRightDash = wb.createCellStyle();
        notificationTableRightDash.cloneStyleFrom(notificationTableHeaderRightDash);
        notificationTableRightDash.setFillPattern(FillPatternType.NO_FILL);
        // ヘッダー以外の両側破線
        notificationTableBothDash = wb.createCellStyle();
        notificationTableBothDash.cloneStyleFrom(notificationTableHeaderBothDash);
        notificationTableBothDash.setFillPattern(FillPatternType.NO_FILL);
        // セルの結合
        sheetNotification.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        sheetNotification.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(2, 2, 2, 4));
        sheetNotification.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(3, 3, 2, 4));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 0, 1));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 3, 5));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 5, 6, 8));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 2, 2));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 9, 9));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 10, 10));
        sheetNotification.addMergedRegion(new CellRangeAddress(5, 6, 11, 11));
        sheetNotification.addMergedRegion(new CellRangeAddress(47, 47, 0, 11));


        /**
         * 行とセルの作成
         * 必要ない。あとの処理でgetじゃなくてcreateを使えば良い
         */
        for (int i = 0; i < rowSizeNotification; i++) {
            row = sheetNotification.createRow(i);
            for (int j = 0; j < colSizeNotification; j++) {
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
        fontNotificationTitleYear.setFontHeightInPoints((short) 16);
        fontNotificationTitleYear.setFontName(PMINCHO);
        styleNotificationTitleYear.setFont(fontNotificationTitleYear);
        styleNotificationTitleYear.setVerticalAlignment(VerticalAlignment.CENTER);
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
        for (int i = 2; i < 4; i++) {
            for (int j = 0; j < 5; j++) {
                sheetNotification.getRow(i).getCell(j).setCellStyle(notificationEmployeeInfoTable);
            }
        }
        // テーブル部分作成
        for (int i = 5; i < 47; i++) {
            // 行番号のループ
            row = sheetNotification.getRow(i);
            for (int j = 0; j < 12; j++) {
                // 列番号のループ
                cell = row.getCell(j);
                if (i == 5) {
                    // テーブル最初の行
                    cell.setCellValue(notificationHeaderStrings[j]);
                    cell.setCellStyle(notificationTableHeaderNorm);
                } else if (i == 6) {
                    // テーブル２番めの行
                    cell.setCellValue(notificationHeaderStrings[j + 12]);
                    // cell.setCellStyle(notificationTableHeaderNorm);
                    switch (j) {
                        case 0:
                        case 3:
                        case 6:
                            cell.setCellStyle(notificationTableHeaderRightDash);
                            break;
                        case 1:
                        case 5:
                        case 8:
                            cell.setCellStyle(notificationTableHeaderLeftDash);
                            break;
                        case 4:
                        case 7:
                            cell.setCellStyle(notificationTableHeaderBothDash);
                            break;
                        default:
                            cell.setCellStyle(notificationTableHeaderNorm);
                    }
                } else {
                    // テーブル３番目以降
                    switch (j) {
                        case 0:
                        case 3:
                        case 6:
                            cell.setCellStyle(notificationTableRightDash);
                            // デバッグ用
                            // System.out.print("i:j" + i + ":" + j);
                            // System.out.print("セルのスタイル：" + cell.getCellStyle().getBorderRight());
                            break;
                        case 1:
                        case 5:
                        case 8:
                            cell.setCellStyle(notificationTableLeftDash);
                            // デバッグ用
                            // System.out.print("i:j" + i + ":" + j);
                            // System.out.print("セルのスタイル：" + cell.getCellStyle().getBorderRight());
                            break;
                        case 4:
                        case 7:
                            cell.setCellStyle(notificationTableBothDash);
                            // デバッグ用
                            // System.out.print("i:j" + i + ":" + j);
                            // System.out.print("セルのスタイル：" + cell.getCellStyle().getBorderRight());
                            break;
                        default:
                            cell.setCellStyle(notificationTableNorm);
                    }
                }

            }

        }
        // デバッグ用出力
        // System.out.println("月のカラムの枠線：" +
        // sheetNotification.getRow(6).getCell(0).getCellStyle().getBorderRight());

        // フッター部分(48行目)
        row = sheetNotification.getRow(47);
        cell = row.getCell(0);
        cell.setCellValue("MicroMagic INC.");
        XSSFCellStyle notificationFooterStyle = wb.createCellStyle();
        XSSFFont notificationFooterFont = wb.createFont();
        notificationFooterFont.setFontName(CENTURY);
        notificationFooterFont.setFontHeightInPoints((short) 16);
        notificationFooterFont.setBold(true);
        notificationFooterFont.setColor(ExcelColor.TODOKEDE_FOOTER);
        notificationFooterStyle.setFont(notificationFooterFont);
        notificationFooterStyle.setAlignment(HorizontalAlignment.CENTER);
        notificationFooterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // notificationFooterStyle.setFillForegroundColor(new XSSFColor(footerRGB, new
        // DefaultIndexedColorMap()));
        cell.setCellStyle(notificationFooterStyle);

        // 項目テーブル
        // セルスタイル
        XSSFCellStyle articleStyle = wb.createCellStyle();
        articleStyle.cloneStyleFrom(notificationEmployeeInfoTable);
        articleStyle.setAlignment(HorizontalAlignment.LEFT);
        XSSFCellStyle articleStyleNum = wb.createCellStyle();
        articleStyleNum.cloneStyleFrom(articleStyle);
        articleStyleNum.setAlignment(HorizontalAlignment.RIGHT);
        // XSSFFont articleFont = wb.createFont();
        XSSFFont articleFont = new XSSFFont();
        articleFont.setFontName(PMINCHO);
        articleFont.setFontHeight(10.5);
        articleStyle.setFont(articleFont);
        XSSFFont articleNumFont = new XSSFFont();
        articleNumFont.setFontName(CENTURY);
        articleStyleNum.setFont(articleNumFont);

        // セルの作成
        for (int i = 5; i < 20; i++) {
            row = sheetNotification.getRow(i);
            for (int j = 14; j < 17; j++) {
                if (i >= 8 && j == 16) {
                    continue;
                }
                cell = row.createCell(j);
                cell.setCellStyle(articleStyle);
                if (j == 14) {
                    cell.setCellValue(i - 4);
                    cell.setCellStyle(articleStyleNum);
                }
                if (j == 15) {
                    cell.setCellValue(articleNoteStrings[i - 5]);
                }
                if (j == 16 && i < 8) {
                    cell.setCellValue(articleNoteStatStrings[i - 5]);
                }

            }
        }
        // カラム幅の指定
        for (int i = 0; i < 17; i++) {
            sheetNotification.setColumnWidth(i, TodokedeWidth.width[i]);
        }
        // 高さの指定
        for (int i = 0; i < 48; i++) {
            if (i < 5) {
                sheetNotification.getRow(i).setHeight(TodokedeHeight.todokedeHeight[i]);
            } else if (i < 47) {
                sheetNotification.getRow(i).setHeight(TodokedeHeight.todokedeHeight[5]);
            } else {
                sheetNotification.getRow(i).setHeight(TodokedeHeight.todokedeHeight[6]);
            }
        }
        // テキストボックス
        XSSFDrawing noteDrawing = sheetNotification.createDrawingPatriarch();
        CreationHelper helper = sheetNotification.getWorkbook().getCreationHelper();
        XSSFClientAnchor anchor = (XSSFClientAnchor) helper.createClientAnchor();
        anchor.setAnchorType(AnchorType.DONT_MOVE_AND_RESIZE);
        anchor.setRow1(36);
        anchor.setDy1(13 * Units.EMU_PER_POINT);
        anchor.setCol1(0);
        anchor.setDx1((int) Math.round(1300 * Units.EMU_PER_PIXEL * Units.DEFAULT_CHARACTER_WIDTH / 256f));
        anchor.setRow2(46);
        anchor.setDy2(5 * Units.EMU_PER_POINT);
        anchor.setCol2(11);
        anchor.setDx2((int) Math.round(6000 * Units.EMU_PER_PIXEL * Units.DEFAULT_CHARACTER_WIDTH / 256f));
        // XSSFClientAnchor anchor = noteDrawing.createAnchor(10 * 256, 10 * 256, 0, 0,
        // (short)2, 2, (short)8, 8);
        XSSFTextBox textBox = noteDrawing.createTextbox(anchor);
        // textBox.setText("注意事項\n\n"
        XSSFRichTextString textString = new XSSFRichTextString();
        String text = "注意事項\n\n" + "･備考欄には、事由、続柄など必要な事項を記入して下さい。" + "\n\n" + "･添付資料が必要な場合は、別途提出して下さい。" + "\n\n"
                + "･事前申請制です。\n" + "　当該シートに記載の上、ファイルごと提出して下さい。\n" + "　止むを得ず事前提出できない場合は、後日速やかに提出して下さい。\n";
        XSSFFont textFont = new XSSFFont();
        textFont.setBold(true);
        textFont.setFontName(PMINCHO);
        textString.append(text, textFont);
        textBox.setText(textString);
        textBox.setLineWidth(1);
        textBox.setLineStyleColor(0, 0, 0);
        textBox.setFillColor(204, 255, 255);
        textBox.setLeftInset(2);
        textBox.setTopInset(2);
        // 影をつけたい


    }



    /**
     * 「振替出勤管理表」シート作成
     */

    void createFurikae(){

        // テーブルヘッダー１行目の文字列
        String[] headerStrings = {
            "振替出勤", "", "", "", "", "", 
            "振替休日", "", "", "", "",
        };
        // テーブルヘッダー２行目の文字列
        String[] headerStrings2 = {
            "月", "日", "曜日", "時間", "備考(事由等)", "",
            "月", "日", "曜日", "日数", "状況"
        };
        // XSSFSheet sheet = sheetList.get(2);
        XSSFSheet sheetFurikae = sheetList.get(1);
        XSSFCellStyle styleNotificationTitleYear = wb.createCellStyle();
        XSSFFont fontNotificationTitleYear = wb.createFont();
        fontNotificationTitleYear.setBold(true);
        fontNotificationTitleYear.setFontHeightInPoints((short) 16);
        fontNotificationTitleYear.setFontName(PMINCHO);
        styleNotificationTitleYear.setFont(fontNotificationTitleYear);
        styleNotificationTitleYear.setVerticalAlignment(VerticalAlignment.CENTER);

        XSSFCellStyle employTableStyle = wb.createCellStyle();
        employTableStyle.setAlignment(HorizontalAlignment.CENTER);
        employTableStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        employTableStyle.setBorderBottom(BorderStyle.THIN);
        employTableStyle.setBorderTop(BorderStyle.THIN);
        employTableStyle.setBorderLeft(BorderStyle.THIN);
        employTableStyle.setBorderRight(BorderStyle.THIN);
        XSSFFont employTableFont = wb.createFont();
        employTableFont.setFontName(PMINCHO);
        employTableStyle.setFont(employTableFont);

        XSSFCellStyle furikaeHeader1 = wb.createCellStyle();
        furikaeHeader1.cloneStyleFrom(notificationEmployeeInfoTable);
        furikaeHeader1.setAlignment(HorizontalAlignment.CENTER);
        furikaeHeader1.setVerticalAlignment(VerticalAlignment.CENTER);
        furikaeHeader1.setFillForegroundColor(ExcelColor.FURIKAE_PINK);
        furikaeHeader1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont furikaeHeader1Font = wb.createFont();
        furikaeHeader1Font.setFontName(PMINCHO);
        furikaeHeader1Font.setColor(ExcelColor.FURIKAE_WHITE);
        furikaeHeader1Font.setBold(true);
        furikaeHeader1.setFont(furikaeHeader1Font);
        
        XSSFCellStyle furikaeHeader2 = wb.createCellStyle();
        furikaeHeader2.cloneStyleFrom(notificationEmployeeInfoTable);
        furikaeHeader2.setAlignment(HorizontalAlignment.CENTER);
        furikaeHeader2.setVerticalAlignment(VerticalAlignment.CENTER);
        furikaeHeader2.setFillForegroundColor(ExcelColor.FURIKAE_BLUE);
        furikaeHeader2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFFont furikaeHeader2Font = wb.createFont();
        furikaeHeader1Font.setFontName(PMINCHO);
        furikaeHeader2Font.setColor(ExcelColor.FURIKAE_WHITE);
        furikaeHeader2Font.setBold(true);
        furikaeHeader2.setFont(furikaeHeader2Font);


        // セル結合する
        sheetFurikae.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
        sheetFurikae.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
        sheetFurikae.addMergedRegion(new CellRangeAddress(2, 2, 2, 4));
        sheetFurikae.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
        sheetFurikae.addMergedRegion(new CellRangeAddress(3, 3, 2, 4));
        sheetFurikae.addMergedRegion(new CellRangeAddress(5, 5, 0, 5));
        sheetFurikae.addMergedRegion(new CellRangeAddress(5, 5, 6, 10));
        sheetFurikae.addMergedRegion(new CellRangeAddress(6, 6, 4, 5));
        for (int i = 7; i < 52; i++) {
            sheetFurikae.addMergedRegion(new CellRangeAddress(i, i, 4, 5));
        }
        sheetFurikae.addMergedRegion(new CellRangeAddress(52, 52, 0, 10));

        // カラム幅の指定
        for (int i = 0; i < 16; i++) {
            sheetFurikae.setColumnWidth(i, FurikaeWidth.width[i]);
        }
        // 高さの指定
        for (int i = 0; i < 53; i++) {
            sheetFurikae.createRow(i).setHeight(FurikaeHeight.height[i]);
        }

        // YYYY年度振替出勤管理表
        cell = sheetFurikae.getRow(0).createCell(0);
        cell.setCellValue("2021年度振替出勤管理表");
        cell.setCellStyle(styleNotificationTitleYear);

        // 社員テーブル
        for(int i = 2; i < 4; i++){
            row = sheetFurikae.getRow(i);
            for(int j = 0; j < 5; j++){
                cell = row.createCell(j);
                cell.setCellStyle(notificationEmployeeInfoTable);
                if(i == 2){
                    if(j == 0) cell.setCellValue("社員番号");
                    if(j == 2) cell.setCellValue("氏名");
                }
                if(i == 3){
                    if(j == 0) cell.setCellValue(employeeCode);
                    if(j == 2) cell.setCellValue(name);
                }
            }
        }

        // 振替テーブル
        // ヘッダー１行目
        row = sheetFurikae.getRow(5);
        for(int i = 0; i < 11; i++){
            cell = row.createCell(i);
            cell.setCellValue(headerStrings[i]);
            if( i < 5){
                cell.setCellStyle(furikaeHeader1);
            } else {
                cell.setCellStyle(furikaeHeader2);
            }

           
        }
        // ヘッダー２行目
        row = sheetFurikae.getRow(6);
        for(int i = 0; i < 11; i++){
            cell = row.createCell(i);
            cell.setCellValue(headerStrings2[i]);
            switch(i) {
                // 右破線
                case 0: case 6:
                    cell.setCellStyle(notificationTableHeaderRightDash);
                    break;
                // 左破線
                case 5: case 10:
                    cell.setCellStyle(notificationTableHeaderLeftDash);
                    break;
                // 両破線
                default:
                    cell.setCellStyle(notificationTableHeaderBothDash);
                    break;
            }

        }
        // 入力テーブル
        for(int i = 7; i < 52; i++){
            row = sheetFurikae.getRow(i);
            for(int j = 0; j < 11; j++){
                cell = row.createCell(j);
                switch(j) {
                    case 0: case 6:
                        cell.setCellStyle(notificationTableRightDash);
                        break;
                    case 5: case 10:
                        cell.setCellStyle(notificationTableLeftDash);
                        break;
                    default:
                        cell.setCellStyle(notificationTableBothDash);
                        break;
                }
            }
        }

        // 凡例テーブル
        for(int i = 5; i < 13; i++){
            row = sheetFurikae.getRow(i);
            for(int j = 13; j < 16; j++){
                cell = row.createCell(j);


            }
        }

        // フッター部分(53行目)
        row = sheetFurikae.getRow(52);
        cell = row.createCell(0);
        cell.setCellValue("MicroMagic INC.");
        XSSFCellStyle notificationFooterStyle = wb.createCellStyle();
        notificationFooterStyle.setAlignment(HorizontalAlignment.CENTER);
        notificationFooterStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        XSSFFont notificationFooterFont = wb.createFont();
        notificationFooterFont.setFontName(CENTURY);
        notificationFooterFont.setFontHeightInPoints((short) 16);
        notificationFooterFont.setBold(true);
        notificationFooterFont.setColor(ExcelColor.TODOKEDE_FOOTER);
        notificationFooterStyle.setFont(notificationFooterFont);
        // notificationFooterStyle.setFillForegroundColor(new XSSFColor(footerRGB, new
        // DefaultIndexedColorMap()));
        cell.setCellStyle(notificationFooterStyle);




    }


    /**
     * 月ごとのシート作成
     */
    void createMonthly() {

        /**
         * テーブルヘッダー文言
         */
        String[] tableHeaderStrings = { 
            "日付", "", "作業項目", "備考", 
            "開始時間", "終了時間", "全時間", "作業時間", "残業時間" };

        //////////////////////////
        // スタイルの定義
        //////////////////////////

        /**
         * 「開始」「締日」スタイル
         */
        CellStyle style4 = wb.createCellStyle();
        Font font4 = wb.createFont();
        font4.setFontHeightInPoints((short) 9);
        font4.setFontName(PMINCHO);
        style4.setFont(font4);
        style4.setVerticalAlignment(VerticalAlignment.CENTER);
        style4.setAlignment(HorizontalAlignment.CENTER);

        /**
         * 開始日・締日フィールドスタイル
         */
        CellStyle style4_2 = wb.createCellStyle();
        Font font4_2 = wb.createFont();
        font4_2.setBold(true);
        font4_2.setUnderline(Font.U_SINGLE);
        font4_2.setFontName(PMINCHO);
        font4_2.setFontHeightInPoints((short) 9);
        style4_2.setFont(font4_2);
        style4_2.setVerticalAlignment(VerticalAlignment.CENTER);

        /**
         * 会社名スタイル
         */
        CellStyle style3 = wb.createCellStyle();
        Font font3 = wb.createFont();
        font3.setFontHeightInPoints((short) 10);
        font3.setFontName(PMINCHO);
        style3.setFont(font3);
        style3.setVerticalAlignment(VerticalAlignment.CENTER);

        /**
         * 「勤怠報告書」スタイル
         */
        CellStyle style2 = wb.createCellStyle();
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font2 = wb.createFont();
        font2.setBold(true);
        font2.setFontHeightInPoints((short) 20);
        font2.setFontName(PMINCHO);
        style2.setFont(font2);

        /**
         * YYYY年MM月分スタイル
         */
        CellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = wb.createFont();
        font.setBold(true);
        font.setFontName(PMINCHO);
        font.setFontHeightInPoints((short) 16);
        style.setFont(font);

        /**
         * 社員テーブルのスタイル
         */
        CellStyle style5 = wb.createCellStyle();
        Font font5 = wb.createFont();
        font5.setFontHeightInPoints((short) 9);
        font5.setFontName(PMINCHO);
        style5.setFont(font5);
        style5.setVerticalAlignment(VerticalAlignment.CENTER);
        style5.setAlignment(HorizontalAlignment.CENTER);
        style5.setBorderBottom(BorderStyle.THIN);
        style5.setBorderTop(BorderStyle.THIN);
        style5.setBorderLeft(BorderStyle.THIN);
        style5.setBorderRight(BorderStyle.THIN);

        /**
         * 提出日のスタイル
         */
        XSSFCellStyle style6 = wb.createCellStyle();
        style6.setAlignment(HorizontalAlignment.RIGHT);
        XSSFFont style6Font = wb.createFont();
        style6Font.setFontName(PMINCHO);
        style6.setFont(style6Font);

        /**
         * 平日のスタイル
         */ 
        XSSFCellStyle styleWeekDay = wb.createCellStyle();
        styleWeekDay.setAlignment(HorizontalAlignment.CENTER);
        styleWeekDay.setFillForegroundColor(ExcelColor.REPORT_BLUE);
        styleWeekDay.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontWeekDay = wb.createFont();
        fontWeekDay.setFontName(PMINCHO);
        fontWeekDay.setFontHeightInPoints((short) 12);
        styleWeekDay.setFont(fontWeekDay);
        // 平日以外のスタイル
        XSSFCellStyle styleWeekEnd = wb.createCellStyle();
        styleWeekEnd.setAlignment(HorizontalAlignment.CENTER);
        styleWeekEnd.setFillForegroundColor(ExcelColor.REPORT_BLUE);
        styleWeekEnd.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontWeekEnd = wb.createFont();
        fontWeekEnd.setColor(IndexedColors.RED.index);
        fontWeekEnd.setFontName(PMINCHO);
        fontWeekEnd.setFontHeightInPoints((short) 12);
        styleWeekEnd.setFont(fontWeekEnd);
        // 右３列のスタイル
        XSSFCellStyle styleThreeRightColumn = wb.createCellStyle();
        styleThreeRightColumn.setAlignment(HorizontalAlignment.CENTER);
        styleThreeRightColumn.setFillForegroundColor(ExcelColor.REPORT_YELLOW);
        styleThreeRightColumn.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styleThreeRightColumn.setFont(fontWeekDay); // フォントは平日と同じ

        /**
         * テーブルヘッダースタイル
         */
        XSSFCellStyle styleTableHeader = wb.createCellStyle();
        styleTableHeader.setAlignment(HorizontalAlignment.CENTER);
        styleTableHeader.setVerticalAlignment(VerticalAlignment.CENTER);
        styleTableHeader.setFillForegroundColor(ExcelColor.REPORT_BLUE);
        styleTableHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font fontTableHeader = wb.createFont();
        fontTableHeader.setBold(true);
        fontTableHeader.setFontHeightInPoints((short) 9);
        fontTableHeader.setFontName(PMINCHO);
        styleTableHeader.setFont(fontTableHeader);


        /**
         * 合計のスタイル
         */
        XSSFCellStyle styleGoukei = wb.createCellStyle();
        styleGoukei.cloneStyleFrom(styleTableHeader);
        styleGoukei.setBorderTop(BorderStyle.DOUBLE);
        styleGoukei.setBorderBottom(BorderStyle.THICK);
        // 端以外の合計行のスタイル
        XSSFCellStyle styleTotalNotEdge = wb.createCellStyle();
        styleTotalNotEdge.cloneStyleFrom(styleGoukei);


        ////////////////////////
        // １年分のシートを作成
        ////////////////////////

        DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("yyyy年M日d日");
        // スタート日付を４月１日とする
        LocalDate startYearDate = LocalDate.of(year, 4, 1);

        // １年分のシートを作成
        for(int i = 0; i < 12; i++) {
            LocalDate startDate = startYearDate.plusMonths(i);
            int month = startDate.getMonthValue();
            int date = startDate.getDayOfMonth();
            sheetList.add(wb.createSheet(month + "月"));
        }
        

        /////////////////////
        // 以下１シート分
        // できればメソッド化して共通化したい
        /////////////////////

        XSSFSheet sheet;

        
        for(int j = 0; j < 12; j++){
            LocalDate startDate = startYearDate.plusMonths(j);
            sheet = wb.getSheet(startDate.getMonthValue() + "月");
            // セル結合している部分は最初に結合する
            // sheet.getRow(0).getCell(0).setCellValue("2021年" + j + "月分");
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
            sheet.addMergedRegion(new CellRangeAddress(5, 5, 6, 8));
            sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, 1));
            sheet.addMergedRegion(new CellRangeAddress(38, 38, 0, 1));
            
            // 幅
            int[] width = {
                1962, 1962, 7082, 4437, 2474, 
                2474, 2474, 2474, 2474, 3640, 
                2417, 2417, 2417, 2417, 2304, 
                2304, 2304
            };

            short[] height = {
                468, 300, 300, 300, 870, 420, 450, 450, 450, 450, 
                450, 450, 450, 450, 450, 450, 450, 450, 450, 450, 
                450, 450, 450, 450, 450, 450, 450, 450, 450, 450, 
                450, 450, 450, 450, 450, 450, 450, 450, 498, 315, 408 
            };

            // 列幅を設定する
            for(int i = 0; i < width.length; i++){
                sheet.setColumnWidth(i, width[i]);
            }

            // 高さを設定する
            for(int i = 0; i < height.length; i++){
                row = sheet.createRow(i);
                row.setHeight(height[i]);
            }

            
            // sheet = wb.getSheet("4月");
            // 「YYYY年M月分」
            // merge cells of 年月
            // 引数の番号は0基底
    
            // フォントの設定は一括で行い、セルに対してセットする
            // セルごとにフォントをセットしないといけない？
            
            // CellStyle はセルごとにインスタンスを生成する必要がある
            row = sheet.getRow(0);
            cell = row.createCell(0);
            cell.setCellValue(startDate.getYear() + "年" + startDate.getMonthValue() + "月分");
            cell.setCellStyle(style);
    
            // 「勤怠報告書」
            cell = row.createCell(3);
            cell.setCellValue("勤怠報告書");
            cell.setCellStyle(style2);
    
            // 会社名の設定
            cell = row.createCell(7);
            cell.setCellValue("株式会社マイクロマジック");
    
            cell.setCellStyle(style3);
    
            // 「開始」
            row = sheet.getRow(1);
            cell = row.createCell(1);
            cell.setCellValue("開始");
    
            cell.setCellStyle(style4);
            // 「YYYY年M月D日」
            cell = row.createCell(2);
            cell.setCellValue(startYearDate.plusMonths(j).format(dateFormat));
            cell.setCellStyle(style4_2);
    
            // 「締日」
            row = sheet.getRow(2);
            cell = row.createCell(1);
            cell.setCellValue("締日");
            cell.setCellStyle(style4);
            // 「YYYY年M月D日」
            cell = row.createCell(2);
            cell.setCellValue(startDate.plusMonths(1).minusDays(1).format(dateFormat));
            cell.setCellStyle(style4_2);
    
            // 「社員番号」
            row = sheet.getRow(3);
            cell = row.createCell(1);
            cell.setCellValue("社員番号");
    
            cell.setCellStyle(style5);
            // 「氏名」
            cell = row.createCell(2);
            cell.setCellValue("氏名");
            cell.setCellStyle(style5);
            // 「担当」
            cell = row.createCell(7);
            cell.setCellValue("担当");
            cell.setCellStyle(style5);
            // 「確認」
            cell = row.createCell(8);
            cell.setCellValue("確認");
            cell.setCellStyle(style5);
    
            // 社員番号の値
            row = sheet.getRow(4);
            cell = row.createCell(1);
            cell.setCellValue(employeeCode);
            // セルスタイルとフォントはstyle5,font5と共通
            cell.setCellStyle(style5);
            // 氏名の値
            cell = row.createCell(2);
            cell.setCellValue(name);
            cell.setCellStyle(style5);
            // 担当欄
            cell = row.createCell(7);
            cell.setCellStyle(style5);
            // 確認欄
            cell = row.createCell(8);
            cell.setCellStyle(style5);
    
            // 「提出日」
            row = sheet.getRow(5);
            cell = row.createCell(6);
            cell.setCellValue("提出日 " + startDate.plusMonths(1).minusDays(1).format(dateFormat));
            

            cell.setCellStyle(style6);
    
    
            // テーブルヘッダー
            row = sheet.getRow(6);
    
            // 「日付」
            // 「作業項目」「備考」「開始時間」「終了時間」「全時間」「作業時間」「残業時間」
            for (int i = 0; i < 9; i++) {
                cell = row.createCell(i);
                cell.setCellValue(tableHeaderStrings[i]);
                cell.setCellStyle(styleTableHeader);
            }
    
            // for (int i = 2; i < 9; i++) {
            //     cell = row.createCell(i);
            //     cell.setCellStyle(styleTableHeader);
            // }
    
            // カレンダー作成
            // List<XSSFCell> cellList = new ArrayList<>();
            // XSSFCell[] cellB = null;
            XSSFCell cellA = null;
            XSSFCell cellB = null;
            
            LocalDate date;
            for (int i = 0; i < 31; i++) {
                row = sheet.getRow(i + 7);
                cellA = row.createCell(0);
                cellB = row.createCell(1);
    
                // 右３列にスタイルを適用
                row.createCell(6).setCellStyle(styleThreeRightColumn);
                row.createCell(7).setCellStyle(styleThreeRightColumn);
                row.createCell(8).setCellStyle(styleThreeRightColumn);
    
                // 日付を加算する
                date = startDate.plusDays(i);
                cellA.setCellValue(date.getDayOfMonth());
                // 日付を
                cellB.setCellValue(date.getDayOfWeek().getDisplayName(TextStyle.SHORT, Locale.JAPANESE));
    
                // セルスタイルのセット
                if (date.getDayOfWeek() == DayOfWeek.SATURDAY || date.getDayOfWeek() == DayOfWeek.SUNDAY) {
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
            XSSFCell totalCell = row.createCell(0);
            totalCell.setCellValue("合計");

            totalCell.setCellStyle(styleGoukei);
            row.createCell(1).setCellStyle(styleGoukei);
    
            // styleTotalNotEdge.setBorderTop(BorderStyle.DOUBLE);
            // styleTotalNotEdge.setBorderBottom(BorderStyle.THICK);
    
            for (int i = 2; i < 9; i++) {
                row.createCell(i).setCellStyle(styleTotalNotEdge);
            }


        }
        

    }


}
