package input;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

import ExcelUtil.ExcelUtil;
import com.alibaba.fastjson.JSON;
import org.apache.commons.logging.Log;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ExcelReader {
    private static Logger LG = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 根据excal路径生成实体集合
     *
     * @param filePath
     * @return
     * @author Changhai
     * @data 2017-7-5
     */
    public static String getJsonString(String filePath) {
        InputStream is;
        try {
            is = new FileInputStream(filePath);
            return getList(is);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 根据输入流生成实体集合
     *
     * @param is
     * @return
     * @throws IOException
     * @author Changhai
     * @data 2017-7-5
     */
    public static String getList(InputStream is)
            throws IOException {
        List<List<String>> list = ExcelReader.readExcel(is);

        //-----------------------遍历数据到实体集合开始-----------------------------------
        List<JSONObject> listBean = new ArrayList<JSONObject>();
        for (int i = 1; i < list.size(); i++) {// i=1是因为第一行不要
            JSONObject obj = new JSONObject();
            List<String> listStr = list.get(i);
            if (listStr.get(1).equals("KEY_3")) {
                System.out.println(listStr.get(4));
                System.out.println(listStr.get(5));
            }
            obj.put("key", listStr.get(1));
            obj.put("value_zh", listStr.get(2));
            obj.put("value_en", listStr.get(3));
            obj.put("value_myz", listStr.get(4));
            obj.put("value_my", listStr.get(5));
            listBean.add(obj);
        }
        //----------------------------遍历数据到实体集合结束----------------------------------
        return JSON.toJSONString(listBean);
    }

    /**
     * Excel读取 操作
     */
    public static List<List<String>> readExcel(InputStream is)
            throws IOException {
        Workbook wb = null;
        try {
            wb = WorkbookFactory.create(is);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        /** 得到第一个sheet */
        Sheet sheet = wb.getSheetAt(0);
        /** 得到Excel的行数 */
        int totalRows = sheet.getPhysicalNumberOfRows();

        /** 得到Excel的列数 */
        int totalCells = 0;
        if (totalRows >= 1 && sheet.getRow(0) != null) {
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
        }

        List<List<String>> dataLst = new ArrayList<List<String>>();
        /** 循环Excel的行 */
        for (int r = 0; r < totalRows; r++) {
            Row row = sheet.getRow(r);
            if (row == null)
                continue;
            List<String> rowLst = new ArrayList<String>();
            /** 循环Excel的列 ,第0列为表的行数，获取列数的时候获取不到，所以获取真实数据的时候+1列才能获取所有真实数据*/
            for (int c = 0; c <= totalCells; c++) {
                Cell cell = row.getCell(c);
                String cellValue = "";
                if (null != cell) {
                    HSSFDataFormatter hSSFDataFormatter = new HSSFDataFormatter();
                    cellValue = hSSFDataFormatter.formatCellValue(cell);

// 以下是判断数据的类型
/*

                    switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC: // 数字
                        cellValue = cell.getNumericCellValue() + "";
                        break;
                    case Cell.CELL_TYPE_STRING: // 字符串
                        cellValue = cell.getStringCellValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN: // Boolean
                        cellValue = cell.getBooleanCellValue() + "";
                        break;
                    case Cell.CELL_TYPE_FORMULA: // 公式
                        cellValue = cell.getCellFormula() + "";
                        break;
                    case Cell.CELL_TYPE_BLANK: // 空值
                        cellValue = "";
                        break;
                    case Cell.CELL_TYPE_ERROR: // 故障
                        cellValue = "非法字符";
                        break;
                    default:
                        cellValue = "未知类型";
                        break;
                    }*/
                }
                rowLst.add(cellValue);
            }
            /** 保存第r行的第c列 */
            dataLst.add(rowLst);
        }
        return dataLst;
    }

    public static void main(String[] args) {
        // TODO Auto-generated method stub
        try {
            //根据流
//            InputStream is = new FileInputStream("d:\\doc\\zegopay.v12.xlsx");
//            String jsonString = ExcelReader.getList(is);
            String jsonString = getJsonString("d:\\doc\\zegopay.v12.xlsx");
            //根据文件路径
            //List<User> list = (List<User>) ExcelReader.getList("d:\\user.xlsx");
            createJsonFile(jsonString, "d:\\doc", "service_message");
            System.out.println("JsonString:" + jsonString);
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    /**
     * 生成.json格式文件
     */
    public static boolean createJsonFile(String jsonString, String filePath, String fileName) {
        // 标记文件生成是否成功
        boolean flag = true;

        // 拼接文件完整路径
        String fullPath = filePath + File.separator + fileName + ".json";

        // 生成json格式文件
        try {
            // 保证创建一个新文件
            File file = new File(fullPath);
            if (!file.getParentFile().exists()) { // 如果父目录不存在，创建父目录
                file.getParentFile().mkdirs();
            }
            if (file.exists()) { // 如果已存在,删除旧文件
                file.delete();
            }
            file.createNewFile();

            if (jsonString.indexOf("'") != -1) {
                //将单引号转义一下，因为JSON串中的字符串类型可以单引号引起来的
                jsonString = jsonString.replaceAll("'", "\\'");
            }
            if (jsonString.indexOf("\"") != -1) {
                //将双引号转义一下，因为JSON串中的字符串类型可以单引号引起来的
                jsonString = jsonString.replaceAll("\"", "\\\"");
            }

            if (jsonString.indexOf("\r\n") != -1) {
                //将回车换行转换一下，因为JSON串中字符串不能出现显式的回车换行
                jsonString = jsonString.replaceAll("\r\n", "\\u000d\\u000a");
            }
            if (jsonString.indexOf("\n") != -1) {
                //将换行转换一下，因为JSON串中字符串不能出现显式的换行
                jsonString = jsonString.replaceAll("\n", "\\u000a");
            }

            // 格式化json字符串
            jsonString = JsonFormatTool.formatJson(jsonString);

            // 将格式化后的字符串写入文件
            Writer write = new OutputStreamWriter(new FileOutputStream(file), "UTF-8");
            write.write(jsonString);
            write.flush();
            write.close();
        } catch (Exception e) {
            flag = false;
            e.printStackTrace();
        }

        // 返回是否成功的标记
        return flag;
    }
}