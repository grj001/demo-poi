package com.poi.demo.excel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Date;

public class CreateExcelSheet {

    public static void main(String[] args) throws IOException {

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
        Date date = new Date(System.currentTimeMillis());
        String str = sdf.format(date);

        String filePath = "E:\\result\\" + str + ".xls";//文件路径
        System.out.println("1. 初始化文件路径 : " + filePath + "\n");


        HSSFWorkbook workbook = new HSSFWorkbook(); //创建Excel文件(Workbook)
        System.out.println("2.创建Excel文件\n");

        HSSFSheet sheet = workbook.createSheet(); //创建工作表(Sheet)


        sheet = workbook.createSheet("Test"); //创建工作表(Sheet)
        System.out.println("3.创建工作表sheet Test\n");



        FileOutputStream out = new FileOutputStream(filePath);
        workbook.write(out);//保存Excel文件
        System.out.println("3.报错Excel文件 sheet\n");

        out.close();//关闭文件流
        System.out.println("OK!");
    }

}
