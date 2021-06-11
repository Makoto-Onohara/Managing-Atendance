package com.example.common;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import lombok.Data;

/**
 * 色管理クラス
 */
@Data
public class ExcelColor {

    /**
     * 「届出」シートの色管理
     */
    // フッターの文字色（青）
    public static final byte[] footerRGB = new byte[]{(byte)0, (byte)0, (byte)255};
    public static final XSSFColor footer = new XSSFColor(footerRGB,  new DefaultIndexedColorMap());
    // ヘッダー青っぽい色
    public static final byte[] blueNoteRGB = new byte[]{(byte)204, (byte)255, (byte)255};
    public static final XSSFColor blueNote = new XSSFColor(blueNoteRGB,  new DefaultIndexedColorMap());
    


    /**
     * 「振替出勤管理表」シートの色管理
     */
    


    /**
     * 「作業報告書」シートの色管理
     */
    // ヘッダー青っぽい色
    public static final byte[] blueRGB = new byte[]{(byte)153, (byte)204, (byte)255};
    public static final XSSFColor blue = new XSSFColor(blueRGB,  new DefaultIndexedColorMap());
    // テーブル右黃っぽい色
    public static final byte[] yellowRGB = new byte[]{(byte)255, (byte)255, (byte)153};
    public static final XSSFColor yellow = new XSSFColor(yellowRGB,  new DefaultIndexedColorMap());
    

}