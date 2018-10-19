package com.poi.demo.excel;

import com.sun.org.apache.regexp.internal.RE;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.regex.Pattern;

public class ReadExcel {

    private static Logger log = LoggerFactory.getLogger(ReadExcel.class);

    private static final ReadExcel RE = new ReadExcel();

    /**
     * 使用迭代器进行 read
     * @param path
     * @return
     */
    public static String readXls(String path) {

        log.debug("开始 读取xls格式Excel文件 ...\n");
        StringBuilder text = new StringBuilder("");
        try {
            FileInputStream is = new FileInputStream(path);
            log.debug("开启 FileInputStream 获取Excel路径: " + path + "\n");

            HSSFWorkbook excel = new HSSFWorkbook(is);
            log.debug("获取路径下的Excel文件: " + excel + "\n");

            //获取第一个sheet
            HSSFSheet sheet0 = excel.getSheetAt(0);
            log.debug("获取Excel文件中的第一个 sheet: " + sheet0 + "\n");

            log.debug("开始遍历sheet .. \n");
            int lineNumber = 0;
            for (Iterator rowIterator = sheet0.iterator(); rowIterator.hasNext(); ) {
                log.debug("开始遍历 sheet 中的第" + lineNumber + " 行 中的 单元格---\n");


                HSSFRow row = (HSSFRow) rowIterator.next();

                int cellNumber = 0;
                for (Iterator iterator = row.cellIterator(); iterator.hasNext(); ) {
                    log.debug("开始遍历 sheet 中的第" + lineNumber + " 行 中的 第" + cellNumber + " 个单元格\n");

                    HSSFCell cell = (HSSFCell) iterator.next();
                    //根据单元的的类型 读取相应的结果
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                         text.append(cell.getStringCellValue() + "\t");
                    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                         text.append(cell.getNumericCellValue() + "\t");
                    } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                         text.append(cell.getCellFormula() + "\t");
                    }

                    cellNumber++;
                }
                 text.append("\n");
                lineNumber++;
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.warn(e.toString());
        }

        log.debug("读取单元格结束...\n");
        return text.toString();
    }

    public static String readXlsx(String path) {

        log.debug("开始 读取xlsx格式Excel文件 ...\n");

        StringBuilder text = new StringBuilder("");
        try {
            OPCPackage pkg = OPCPackage.open(path);
            log.debug("开启 OPCPackage 获取Excel路径: " + path + "\n");

            XSSFWorkbook excel = new XSSFWorkbook(pkg);
            log.debug("获取路径下的Excel文件: " + excel + "\n");

            //获取第一个sheet
            XSSFSheet sheet0 = excel.getSheetAt(0);
            log.debug("获取Excel文件中的第一个 sheet: " + sheet0 + "\n");


            log.debug("开始遍历sheet .. \n");
            int lineNumber = 0;
            for (Iterator rowIterator = sheet0.iterator(); rowIterator.hasNext(); ) {
                log.debug("开始遍历 sheet 中的第" + lineNumber + " 行 中的 单元格---\n");

                XSSFRow row = (XSSFRow) rowIterator.next();

                int cellNumber = 0;
                for (Iterator iterator = row.cellIterator(); iterator.hasNext(); ) {
                    log.debug("开始遍历 sheet 中的第" + lineNumber + " 行 中的 第" + cellNumber + " 个单元格\n");

                    XSSFCell cell = (XSSFCell) iterator.next();
                    //根据单元的的类型 读取相应的结果
                    if (cell.getCellTypeEnum() == CellType.STRING) {
                        text.append(cell.getStringCellValue() + "\t");
                    } else if (cell.getCellTypeEnum() == CellType.NUMERIC) {
                        text.append(cell.getNumericCellValue() + "\t");
                    } else if (cell.getCellTypeEnum() == CellType.FORMULA) {
                        text.append(cell.getCellFormula() + "\t");
                    }
                }
                text.append("\n");
            }
        } catch (Exception e) {
            e.printStackTrace();
            log.warn(e.toString());
        }

        log.debug("读取单元格结束...\n");
        return text.toString();
    }

    //判断是否是姓名
    public boolean isName(String name){
        return Pattern.matches("[\\u4e00-\\u9fa5]{2,3}",name);
    }
    //判断是否是数字
    public boolean isNumber(String number){
        return Pattern.matches("[\\d]+[.]?[\\d]*",number);
    }

    public static void readOutputExcelOnceEmp(){
        System.out.println("---readOutputExcelOnceEmp begin");



        boolean isName = true;
        ArrayList<HashMap<String,Float>> lm = new ArrayList<HashMap<String,Float>>();

        HSSFWorkbook excel = null;
        try {
            FileInputStream is = new FileInputStream("C:\\Users\\HSF2015\\Desktop\\workData.xls");
            System.out.println("开启 FileInputStream 获取Excel路径: " + "C:\\Users\\HSF2015\\Desktop\\workData.xls" + "\n");

            excel = new HSSFWorkbook(is);
            System.out.println("获取路径下的Excel文件: " + excel + "\n");
        } catch (IOException e) {
            e.printStackTrace();
        }


        //获取第一个sheet
        HSSFSheet sheet0 = excel.getSheetAt(0);
        System.out.println("获取Excel文件中的第一个 sheet: " + sheet0 + "\n");

        for(int i=0;i<sheet0.getLastRowNum();i++){

            HSSFRow row = sheet0.getRow(i);


            //判断第一列是否是姓名
            if(RE.isName(sheet0.getRow(i).getCell(0).toString()) && i+1>=5){
                System.out.print("第"+i+"行 是姓名行\t");


                String name = "";
                Float number = 0F;

                for(int j=0;j<row.getLastCellNum();j++){

                    HSSFCell cell = row.getCell(j);
                    String cellValue = cell.toString().trim();


                    if(RE.isName(cellValue)){
                        System.out.print(cellValue+"\t");
                        name = cellValue;
                    }else if(RE.isNumber(cellValue)){
                        System.out.print(cellValue+"\t");
                        number = Float.valueOf(cellValue);
                    }


                }
                HashMap map = new HashMap<String,Float>();
                map.put(name,number);
                lm.add(map);
                System.out.println();

            }else{
                System.out.println("第"+i+"行 不是姓名行\t");
                continue;
            }
        }

        System.out.println(lm);

        System.out.println("---readOutputExcelOnceEmp end");
    }

    public static void main(String[] args) {
        readOutputExcelOnceEmp();
    }

}
