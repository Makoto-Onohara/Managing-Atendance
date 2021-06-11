package com.example.common;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import lombok.Data;

/**
 * 色管理クラス
 */
@Data
public final class ExcelColor {

    /**
     * newできないようにコンストラクタをprivateにする
     */
    private ExcelColor(){}

    /**
     * 「届出」シートの色管理
     */
    // フッターの文字色（青）
    public static final byte[] footerRGB = new byte[]{(byte)0, (byte)0, (byte)255};
    public static final XSSFColor TODOKEDE_FOOTER = new XSSFColor(footerRGB,  new DefaultIndexedColorMap());
    // ヘッダー青っぽい色
    public static final byte[] blueNoteRGB = new byte[]{(byte)204, (byte)255, (byte)255};
    public static final XSSFColor BLUE_NOTE = new XSSFColor(blueNoteRGB,  new DefaultIndexedColorMap());
    


    /**
     * 「振替出勤管理表」シートの色管理
     */
    // ヘッダー「振替出勤」
    public static final byte[] furikaePink = new byte[]{(byte)255, (byte)0, (byte)255};
    public static final XSSFColor FURIKAE_PINK = new XSSFColor(furikaePink, new DefaultIndexedColorMap());
    // ヘッダー「振替休日」
    public static final byte[] furikaeBlue = new byte[]{(byte)51, (byte)102, (byte)255};
    public static final XSSFColor FURIKAE_BLUE = new XSSFColor(furikaeBlue, new DefaultIndexedColorMap());
    // ヘッダー２列目の薄い青（届出シートの青と同じ）
    public static final XSSFColor FURIKAE_LIGHT_BLUE = BLUE_NOTE;
    


    /**
     * 「作業報告書」シートの色管理
     */
    // ヘッダー青っぽい色
    public static final byte[] blueRGB = new byte[]{(byte)153, (byte)204, (byte)255};
    public static final XSSFColor REPORT_BLUE = new XSSFColor(blueRGB,  new DefaultIndexedColorMap());
    // テーブル右黃っぽい色
    public static final byte[] yellowRGB = new byte[]{(byte)255, (byte)255, (byte)153};
    public static final XSSFColor REPORT_YELLOW = new XSSFColor(yellowRGB,  new DefaultIndexedColorMap());
    

}