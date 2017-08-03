package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class ExcelTool {

    public void testExcel() throws Exception {
        FileInputStream fis = new FileInputStream(new File("/Users/user/Desktop/py/excel/Test.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        XSSFRow row1 = sheet.getRow(1);
        XSSFRow row2 = sheet.getRow(2);

        ArrayList<XSSFRow> topRow = new ArrayList<>();
        row.setRowNum(0);
        topRow.add(row);
        topRow.add(row1);
        topRow.add(row2);

        CellCopyPolicy cellP = new CellCopyPolicy();
        cellP.setCopyCellStyle(false);
        cellP.setCondenseRows(true);
        FileOutputStream out = new FileOutputStream(new File("/Users/user/Desktop/py/excel/Test1.xlsx"));
        FileInputStream in = new FileInputStream(new File("/Users/user/Desktop/py/excel/Test1.xlsx"));

        //创建新工作簿
        XSSFWorkbook book = new XSSFWorkbook(in);
        XSSFSheet sheets = book.createSheet("1");
        sheets.copyRows(topRow, 0, cellP);
        book.write(out);
        book.close();
        out.close();
    }

    public void cutExcel(String basePath, String filename, String outPutPath,
                         int isWindows, int collectionCol, int sheetIndex,
                         int cutLength, int totalRows) throws Exception {
        String path = basePath + (isWindows == 1 ? "\\" : "/") + filename;
        FileInputStream fis = new FileInputStream(new File(path));
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
        Map<String, Integer> ros = new HashMap<>();//缓存名称和具体的行数
        Map<String, Integer> preRos = new HashMap<>();//缓存处理之前的行数
        Map<String, FileInputStream> tarInFiless = new HashMap<>();//缓存
        Map<String, XSSFWorkbook> works = new HashMap<>();
        Map<String, FileOutputStream> tarOutFiles = new HashMap<>();
        Map<String, ArrayList<XSSFRow>> targetRows = new HashMap<>();
        XSSFRow top = sheet.getRow(0);//表头
        top.setRowNum(0);
        int len = totalRows - 1;
        int restSize = len;
        int startIndex = 1;
        System.out.println("该文件总行数： " + len);
        //配置信息
        ArrayList<XSSFRow> topRow = new ArrayList<>();
        topRow.add(top);
        CellCopyPolicy cellP = new CellCopyPolicy();
        cellP.setCopyCellStyle(false);
        cellP.setCondenseRows(true);
        int curSize = 0;
        while (restSize > 0) {
            if (restSize > cutLength) {
                curSize = cutLength;
                restSize -= cutLength;
            } else {
                curSize = restSize;
                restSize = 0;
            }
            for (int i = 0; i < curSize; i++) {
                XSSFRow row = sheet.getRow(startIndex + i);
                if (row == null) {
                    System.out.println("This is a null row " + (startIndex + i));
                    continue;
                }
                String val = row.getCell(collectionCol).getStringCellValue();
                if (val == null || val.equals("")) continue;
                if (ros.containsKey(val)) {
                    //缓存这个
                    row.setRowNum(ros.get(val) + 1);
                    if (targetRows.containsKey(val)) {
                        targetRows.get(val).add(row);
                    } else {
                        ArrayList<XSSFRow> newRow = new ArrayList<>();
                        row.setRowNum(1);
                        newRow.add(row);
                        targetRows.put(val, newRow);
                    }
                } else {
                    System.out.println("New Area: " + val);
                    ArrayList<XSSFRow> newRow = new ArrayList<>();
                    row.setRowNum(1);
                    newRow.add(row);
                    targetRows.put(val, newRow);
                    ros.put(val, 0);
                    preRos.put(val, 1);
                }
            }
            startIndex += curSize;

            for (String curKey : targetRows.keySet()) {
                File tarF = new File(outPutPath + (isWindows == 1 ? "\\" : "/") + curKey + ".xlsx");
                if (tarF.exists()) {
                    FileInputStream tarIn = new FileInputStream(tarF);
//                    if (tarInFiless.containsKey(curKey)) {
//                        tarIn = tarInFiless.get(curKey);
//                    } else {
//                        tarIn = new FileInputStream(tarF);
//                        tarInFiless.put(curKey, tarIn);
//                    }
                    //创建新工作簿
                    XSSFWorkbook book = new XSSFWorkbook(tarIn);
//                    if (works.containsKey(curKey)) {
//                        book = works.get(curKey);
//                    } else {
//                        book = new XSSFWorkbook(tarIn);
//                        works.put(curKey, book);
//                    }
                    XSSFSheet sheets = book.getSheetAt(0);
                    sheets.copyRows(targetRows.get(curKey), ros.get(curKey) + 1, cellP);
                    ros.put(curKey, ros.get(curKey) + targetRows.get(curKey).size());
//                    FileOutputStream out = null;
//                    if (tarOutFiles.containsKey(curKey)) {
//                        out = tarOutFiles.get(curKey);
//                    } else {
//                        out = new FileOutputStream(tarF);
//                        tarOutFiles.put(curKey, out);
//                    }
                    FileOutputStream out = new FileOutputStream(tarF);
//                    System.out.println("The area " + curKey + " : " + ros.get(curKey));
                    book.write(out);
                    book.close();
                    out.close();
                    tarIn.close();
                } else {
                    //创建新工作簿
                    XSSFWorkbook book = new XSSFWorkbook();
                    XSSFSheet sheets = book.createSheet("1");
                    sheets.copyRows(topRow, 1, cellP);
                    sheets.copyRows(targetRows.get(curKey), 2, cellP);
                    ros.put(curKey, ros.get(curKey) + targetRows.get(curKey).size());
//                    FileOutputStream out = null;
//                    if (tarOutFiles.containsKey(curKey)) {
//                        out = tarOutFiles.get(curKey);
//                    } else {
//                        out = new FileOutputStream(tarF);
//                        tarOutFiles.put(curKey, out);
//                    }
                    FileOutputStream out = new FileOutputStream(tarF);
//                    System.out.println("The area " + curKey + " : " + ros.get(curKey));
                    book.write(out);
                    book.close();
                    out.close();
                }
//                FileOutputStream out = new FileOutputStream(tarF);
//                book.write(out);
//                book.close();
//                out.close();
            }
            targetRows.clear();
            System.gc();
            Thread.sleep(10000);
            System.out.println("程序执行进度:" + ((len - restSize) * 100) / len + "%");
        }
        int count = 0;
        for (String key : ros.keySet()
                ) {
            System.out.println("The area " + key + " : " + ros.get(key));
            count += ros.get(key);
        }
        System.out.println("该程序一共处理行数：" + count);
    }

    public static void main(String[] args) throws Exception {
        ExcelTool excelTool = new ExcelTool();
        CsvTool csvTool = new CsvTool();
//        excelTool.testExcel();
//        Thread.sleep(10000);
//        String base = "/Users/user/Desktop/py/excel";
//        String src = "real.csv";
//        String target = "/Users/user/Desktop/py/excel/target";
        String base = "D:\\temp\\src";
        String src = "src.csv";
        String target = "D:\\temp\\target";
        int isWindows = 1;
        int targetCol = 2;
        int targetSheet = 0;
        int cutLength = 8000;
        int totalRows = 0;
        String temp = "";
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        System.out.println("Excel split(Based on Unique column value) Demo");
        System.out.println("第一个参数：源文件路径");
        System.out.println("默认路径：D:\\temp\\src");
        temp = reader.readLine();
        if (!temp.equals("")) {
            base = temp.trim();
        }
        System.out.println("第二个参数：源文件名称");
        System.out.println("默认名称：src.xlsx");
        temp = reader.readLine();
        if (!temp.equals("")) {
            src = temp.trim();
        }
        System.out.println("第三个参数：处理后目标文件路径");
        System.out.println("默认路径：D:\\temp\\target");
        temp = reader.readLine();
        if (!temp.equals("")) {
            target = temp.trim();
        }
        System.out.println("第四个参数：是否是Windows系统");
        System.out.println("默认名称：是（0:否；1：是）");
        temp = reader.readLine();
        if (!temp.equals("")) {
            isWindows = Integer.parseInt(temp.trim());
        }
        System.out.println("第五个参数：目标分类列");
        System.out.println("默认名称：2（第二列，从1开始）");
        temp = reader.readLine();
        if (!temp.equals("")) {
            targetCol = Integer.parseInt(temp.trim());
        }
        System.out.println("第六个参数：目标Sheet");
        System.out.println("默认名称：0（第一个，从0开始）");
        temp = reader.readLine();
        if (!temp.equals("")) {
            targetSheet = Integer.parseInt(temp.trim());
        }
        System.out.println("第七个参数：切割大小");
        System.out.println("默认名称：8000");
        temp = reader.readLine();
        if (!temp.equals("")) {
            cutLength = Integer.parseInt(temp.trim());
        }

        temp = "";

        while (temp.equals("")) {
            System.out.println("第八个参数：总行数");
            System.out.println("默认名称：必须填写");
            temp = reader.readLine();
            if (!temp.equals("")) {
                totalRows = Integer.parseInt(temp.trim());
            }
        }
//        excelTool.cutExcel(base, src, target, isWindows, targetCol, targetSheet, cutLength, totalRows);
        csvTool.cutExcel(base, src, target, isWindows, targetCol, targetSheet, cutLength, totalRows);
    }
}

