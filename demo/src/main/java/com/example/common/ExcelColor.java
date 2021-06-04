package com.example.common;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import lombok.Data;

@Data
public class ExcelColor {

    // 「届出」シート
    // フッターの文字色（青）
    public static final byte[] footerRGB = new byte[]{(byte)0, (byte)0, (byte)255};
    // ヘッダー青っぽい色
    public static final byte[] blueNoteRGB = new byte[]{(byte)204, (byte)255, (byte)255};
    
    // 「作業報告書」シート
    // ヘッダー青っぽい色
    public static final byte[] blueRGB = new byte[]{(byte)153, (byte)204, (byte)255};
    // テーブル右黃っぽい色
    public static final byte[] yellowRGB = new byte[]{(byte)255, (byte)255, (byte)153};

    public static final XSSFColor blueNote = new XSSFColor(blueNoteRGB,  new DefaultIndexedColorMap());
    public static final XSSFColor blue = new XSSFColor(blueRGB,  new DefaultIndexedColorMap());
    public static final XSSFColor yellow = new XSSFColor(yellowRGB,  new DefaultIndexedColorMap());
    public static final XSSFColor footer = new XSSFColor(footerRGB,  new DefaultIndexedColorMap());

}