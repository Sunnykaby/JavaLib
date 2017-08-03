package excel;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.StringTokenizer;

/**
 * Created by User on 2017/7/4.
 */
public class CsvTool {

    /**
     * 获取index单元格数据
     *
     * @param index   （单元格从1开始）
     * @param srcLine
     * @return
     */
    public String getCell(int index, String srcLine) {
        if (srcLine != null || !srcLine.equals("")) {
            StringTokenizer st = new StringTokenizer(srcLine, ",");
            if (st.hasMoreTokens()) {
                String result = "";
                while (index > 0) {
                    result = st.nextToken();
                    index--;
                }
                return result;
            } else return null;
        } else return null;
    }

    public void cutExcel(String basePath, String filename, String outPutPath,
                         int isWindows, int collectionCol, int sheetIndex,
                         int cutLength, int totalRows) throws Exception {
        String path = basePath + (isWindows == 1 ? "\\" : "/") + filename;
        FileInputStream fis = new FileInputStream(new File(path));
        BufferedReader reader = new BufferedReader(new InputStreamReader(fis, "GBK"));

        System.out.println("该文件总行数： " + totalRows);
        String topLine = reader.readLine();
        Map<String, ArrayList<String>> targetRows = new HashMap<>();
        Map<String, Integer> ros = new HashMap<>();//缓存名称和具体的行数
        int restSize = totalRows;

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
                String row = reader.readLine();
                if (row == null || row.equals("")) {
                    System.out.println("This is a null row");
                    continue;
                }
                String val = getCell(collectionCol, row);
                if (val == null || val.equals("")) continue;
                if (ros.containsKey(val)) {
                    //缓存这个
                    if (targetRows.containsKey(val)) {
                        targetRows.get(val).add(row);
                    } else {
                        ArrayList<String> newRow = new ArrayList<>();
                        newRow.add(row);
                        targetRows.put(val, newRow);
                    }
                } else {
                    System.out.println("New Area: " + val);
                    ArrayList<String> newRow = new ArrayList<>();
                    newRow.add(row);
                    targetRows.put(val, newRow);
                    ros.put(val, 0);
                }
            }

            for (String curKey : targetRows.keySet()) {

                File tarF = new File(outPutPath + (isWindows == 1 ? "\\" : "/") + curKey + ".csv");
                boolean isExis = tarF.exists();
//                FileOutputStream out = new FileOutputStream(tarF, true);
//                OutputStreamWriter outW = new OutputStreamWriter(out, "GBK");
                BufferedWriter wr = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(tarF, true), "GBK"));
//                FileWriter br = new FileWriter(tarF, true);
                if (!isExis) {
                    wr.write(topLine);
                    wr.newLine();
                }
                for (String line : targetRows.get(curKey)
                        ) {
                    wr.write(line);
                    wr.newLine();
                }
                ros.put(curKey, ros.get(curKey) + targetRows.get(curKey).size());
                wr.flush();
                wr.close();
            }
            targetRows.clear();
            System.gc();
            Thread.sleep(10000);
            System.out.println("程序执行进度:" + ((totalRows - restSize) * 100) / totalRows + "%");
        }
        int count = 0;
        for (String key : ros.keySet()
                ) {
            System.out.println("The area " + key + " : " + ros.get(key));
            count += ros.get(key);
        }
        System.out.println("该程序一共处理行数：" + count);
    }

}
