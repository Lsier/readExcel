package input;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
public class ExcelToJson {
    public static void main(String[] args) throws IOException {
        String path="D://doc//zegopay.v12.xlsx";
        System.out.println(readExcel(path));
    }

    /**
     * 读取某一个单元格值
     * @param cell
     * @return
     * @throws Exception
     */
    public static Object getCellValueByCell(Cell cell) throws Exception {
        //判断是否为null或空串
        SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        SimpleDateFormat sdv = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzz yyyy", Locale.US);
        if (cell == null || cell.toString().trim().equals("")) {
            return "";
        }
        Object value = null;
        int cellType = cell.getCellType();
        if (XSSFCell.CELL_TYPE_BLANK == cellType) {
            value = null;
        } else if (XSSFCell.CELL_TYPE_BOOLEAN == cellType) {
            value = cell.getBooleanCellValue();
        } else if (XSSFCell.CELL_TYPE_ERROR == cellType) {
            value = cell.getErrorCellValue();
        } else if (XSSFCell.CELL_TYPE_FORMULA == cellType) {
            value = cell.getNumericCellValue();
        } else if (XSSFCell.CELL_TYPE_NUMERIC == cellType) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = (Date) sdv.parse(cell.getDateCellValue().toString()); //将读出的时间进行格式转化，
                value=sdf.format(date);//最终输出格式为 yyyy-MM-dd HH:mm:ss
            } else {
                value = new DecimalFormat("0.########").format(cell.getNumericCellValue());
            }
        } else if (XSSFCell.CELL_TYPE_STRING == cellType) {
            value = cell.getStringCellValue().trim();
        } else {
            throw new Exception("不能识别类型！");
        }
        return value;
    }

    /**
     * 读取Excel文件转成json
     * @param path
     * @return
     * @throws IOException
     */
    public static String readExcel(String path) throws IOException {
        JSONArray array = new JSONArray();
        InputStream is = null;
        XSSFWorkbook workbook = null;
        try {
            is = new FileInputStream(path);
            workbook = new XSSFWorkbook(is);
            //获得第一个工作表对象(ecxel中sheet的编号从0开始,0,1,2,3,....)
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = sheet.getRow(0);
            for(int i=1;i<=sheet.getLastRowNum();i++){ //从第二行开始读取数据，第一行数据为key,
                JSONObject object=new JSONObject();
                XSSFRow r = sheet.getRow(i); //第2行开始是数据行
                for(int j=0;j<row.getLastCellNum();j++){
                    XSSFCell cell=r.getCell(j);
                    object.put(getCellValueByCell(sheet.getRow(0).getCell(j)).toString(), getCellValueByCell(cell));
                }
                array.add(object);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            workbook.close();
            is.close();
        }
        return array.toString();
    }
}