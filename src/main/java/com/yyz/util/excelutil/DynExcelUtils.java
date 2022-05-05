package com.yyz.util.excelutil;

import cn.hutool.core.io.FileUtil;
import com.yyz.util.constant.Constants;
import com.yyz.util.wordutil.PoiWordUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * excel模板数据替换工具类
 * @Author: yyz
 * @Date: 2022/4/24 16:00
 */
public class DynExcelUtils {

    /**
     * 动态标记行删除标识
     */
    private static boolean deleteFlag = false;

    /**
     * 动态行标记索引位置
     */
    private static Integer rowIndex = null;
    /**
     *
     * @param modelPath
     * @param outPath
     * @param paramsMap
     */
    public void replaceExcelData(String modelPath, String outPath, Map<String,Object> paramsMap){
        replaceExcelData(modelPath,outPath,paramsMap,null);
    }


    private void replaceExcelData(String modelPath, String outPath, Map<String, Object> paramsMap, String pattern) {

        Workbook workbook = null;

        File file = FileUtil.newFile(modelPath);
        if (!file.exists()) {
            System.out.println("模板文件:"+modelPath+"不存在!");
            return;
        }

        try(FileInputStream fileInputStream = new FileInputStream(file)){
            if ("xls".equals(FileUtil.getSuffix(file))) {
                workbook = new HSSFWorkbook(fileInputStream);
            }else if ("xlsx".equals(FileUtil.getSuffix(file))) {
                workbook = new XSSFWorkbook(fileInputStream);
            }else {
                System.out.println("请检查文件扩展名是否为xls或xlsx!");
                return;
            }
            replaceExcelData(paramsMap,outPath,workbook,pattern);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    private void replaceExcelData(Map<String, Object> paramsMap, String outPath, Workbook workbook, String pattern) {

        if (StringUtils.isBlank(pattern)) {
            pattern = "yyyy-MM-dd";
        }

        Sheet sheet = null;

        try {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheet = workbook.getSheetAt(i);
                for (int k = 0; k < sheet.getPhysicalNumberOfRows(); k++) {
                    Row row = sheet.getRow(k);
                    if (row != null) {
                        int num = row.getLastCellNum();
                        for (int j = 0; j < num; j++) {
                            Cell cell =  row.getCell(j);
                            if (cell != null) {
                                cell.setCellType(CellType.STRING);
                            }
                            if (cell == null || cell.getStringCellValue() == null) {
                                continue;
                            }
                            String cellValue = cell.getStringCellValue();
                            if (StringUtils.isNotBlank(cellValue)) {
                                // 与参数map里的key,匹配,则替换表格你的值
                                for (Map.Entry<String, Object> entry : paramsMap.entrySet()) {
                                    String key = entry.getKey();
                                    // 非动态行直接将cell替换为map里的值即可
                                    boolean flag = (Constants.PLACEHOLDER_PREFIX + key + Constants.PLACEHOLDER_SUFFIX).equals(cellValue)
                                            && !cellValue.startsWith(Constants.ADD_ROW_FLAG);
                                    if (flag) {
                                        String value = entry.getValue().toString();
                                        cell.setCellValue(value);
                                        break;
                                    }else if(cellValue.startsWith(Constants.ADD_ROW_FLAG)){
                                        deleteFlag = true;
                                        rowIndex = row.getRowNum();
                                        // 动态行添加
                                        dynRowAdd(sheet,row,cellValue,paramsMap,pattern);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (deleteFlag) {
                // 直接导出的第一行数据有问题,显示不出来,故事后再删除动态标记行
                sheet.shiftRows(rowIndex+1,sheet.getLastRowNum(),-1);
            }
            try(FileOutputStream fileOutputStream = new FileOutputStream(outPath)){
                workbook.write(fileOutputStream);
            }
        }catch (Exception e){
            e.printStackTrace();
        }
    }


    private void dynRowAdd(Sheet sheet, Row row, String cellValue, Map<String, Object> paramsMap, String pattern) {

        List<?> dataList = ( List<?> )paramsMap.get(PoiWordUtils.getKeyFromPlaceholder(cellValue));

        String key = "";

        int index = row.getRowNum();

        for (int i = 0; i < dataList.size(); i++) {
            index ++;
            Row addRow = sheet.createRow(index);
            Object t = dataList.get(i);
            try{
                if (t instanceof Map) {
                    Map<String,Object> map = ( Map<String,Object>)t;
                    int cellNum = 0;
                    // 待添加的数据有多少,添加多少列
                   Iterator<String> it2 = map.keySet().iterator();
                    while (it2.hasNext()) {
                        key = it2.next();
                        Object value= map.get(key);
                        Cell cell = addRow.createCell(cellNum);
                        cellNum = setCellValue(cell,value,pattern,cellNum,null,addRow);
                        cell.getCellStyle().cloneStyleFrom(row.getCell(cellNum).getCellStyle());
                        cellNum ++;
                    }
                }else {
                    List<FieldForSortting> fields = ExcelUtil.sortFieldByAnno(t.getClass());
                    int cellNum = 0;
                    for (int j = 0; j < fields.size(); j++) {
                       Cell cell = addRow.createCell(cellNum);
                        Field field = fields.get(j).getField();
                        field.setAccessible(true);
                        Object value = field.get(t);
                        cellNum = setCellValue(cell,value,pattern,cellNum,field,addRow);
                        cell.getCellStyle().cloneStyleFrom(row.getCell(cellNum).getCellStyle());
                        cellNum++;
                    }
                }
            }catch (Exception e){

            }

        }
    }

    private static int setCellValue(Cell cell,Object value,String pattern,int cellNum,Field field,Row row){
        String textValue = null;
        if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            textValue = sdf.format(date);
        } else if (value instanceof String[]) {
            String[] strArr = (String[]) value;
            for (int j = 0; j < strArr.length; j++) {
                String str = strArr[j];
                cell.setCellValue(str);
                if (j != strArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else if (value instanceof Double[]) {
            Double[] douArr = (Double[]) value;
            for (int j = 0; j < douArr.length; j++) {
                Double val = douArr[j];
                // 值不为空则set Value
                if (val != null) {
                    cell.setCellValue(val);
                }

                if (j != douArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else {
            // 其它数据类型都当作字符串简单处理
            String empty = "";
            if(field != null) {
                ExcelCell anno = field.getAnnotation(ExcelCell.class);
                if (anno != null) {
                    empty = anno.defaultValue();
                }
            }
            textValue = value == null ? empty : value.toString();
        }
        if (textValue != null) {
            if (cell instanceof HSSFCell) {
                RichTextString richString = new HSSFRichTextString(textValue);
                cell.setCellValue(richString);
            }else if (cell instanceof XSSFCell) {
                RichTextString richString = new XSSFRichTextString(textValue);
                cell.setCellValue(richString);
            }else {
                cell.setCellValue(textValue);
            }
        }
        return cellNum;
    }

}
