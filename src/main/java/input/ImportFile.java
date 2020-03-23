package input;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.List;

import com.alibaba.fastjson.JSON;
import org.json.JSONException;
import org.json.JSONObject;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

//import com.siemens.entity.master;
//import com.siemens.service.masterService;
//import com.siemens.serviceImpl.masterServiceImpl;
//import com.siemens.serviceImpl.webServiceImpl;

public class ImportFile {
    public static void main(String[] args) throws JSONException {
//		master masters = new master();
//		ApplicationContext ac = new ClassPathXmlApplicationContext("applicationContext.xml");
//		masterService ms = (masterService)ac.getBean("masterservice");
        Workbook wb = null;
        Sheet sheet = null;
        Row row = null;
        String cellData = null;
        //文件路径，
        String filePath = "d:\\doc\\zegopay.v12.xlsx";
//        String filePath = "d:\\doc\\zegopay12.xls";

        wb = EXCELBean.readExcel(filePath);

        if (wb != null) {
            //用来存放表中数据
            List<JSONObject> listMap = new ArrayList<JSONObject>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            //这里创建json对象，实测用map的话，json数据会有问题
//            JSONObject jsonMap = new JSONObject();
            //循环行
            for (int i = 1; i < rownum; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    //创建list对象接收读出的excel数据
                    List<String> list = new ArrayList<String>();
                    //循环列
                    for (int j = 1; j <= colnum; j++) {
                        cellData = (String) EXCELBean.getCellFormatValue(row.getCell(j));
                        list.add(cellData);
                    }
                    //System.out.println(list.get(59));

                    //下面具体是本人对数据按需求进行格式处理     ---创建json对象会报异常，捕捉一下。
                    JSONObject jsonObject = new JSONObject();
                    if (list.get(0).equals("KEY_3")){
                        String s = list.get(4);
                    }
                    jsonObject.put("key", list.get(0));
                    jsonObject.put("value_zh", list.get(1));
                    jsonObject.put("value_en", list.get(2));
                    jsonObject.put("value_myz", list.get(3));
                    jsonObject.put("value_my", list.get(4));
                    listMap.add(jsonObject);
                } else {
                    break;
                }
            }// end for row
            //最外层加个key-gridData
            String jsonString = JSON.toJSONString(listMap);
            createJsonFile(jsonString, "d:\\doc", "service_message_t");
            System.out.println(jsonString);
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