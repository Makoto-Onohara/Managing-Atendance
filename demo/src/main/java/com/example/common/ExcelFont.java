package com.example.common;

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import lombok.Data;

@Data
public class ExcelFont {


    // ヘッダーの色
    // 青っぽい色
    public static final byte[] blueRGB = new byte[]{(byte)153, (byte)204, (byte)255};
    // 黃っぽい色
    public static final byte[] yellowRGB = new byte[]{(byte)255, (byte)255, (byte)153};
    // フッターの文字色（青）
    public static final byte[] footerRGB = new byte[]{(byte)0, (byte)0, (byte)255};

    public static final XSSFColor blue = new XSSFColor(blueRGB,  new DefaultIndexedColorMap());
    public static final XSSFColor yellow = new XSSFColor(yellowRGB,  new DefaultIndexedColorMap());
    public static final XSSFColor footer = new XSSFColor(footerRGB,  new DefaultIndexedColorMap());

}