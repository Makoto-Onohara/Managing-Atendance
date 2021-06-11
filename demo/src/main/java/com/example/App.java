package com.example;

import java.io.IOException;

import com.example.Excel.CreateExcel;

/**
 * Hello world!
 */
public class App 
{
    public static void main( String[] args ) throws IOException
    {

        CreateExcel excel = new CreateExcel();
        // エクセル作成の実行
        excel.createExcel();
    }
}
