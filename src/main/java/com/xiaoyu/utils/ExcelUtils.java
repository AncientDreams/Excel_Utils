package com.xiaoyu.utils;

import com.xiaoyu.anno.ExcelField;
import com.xiaoyu.anno.FiledValue;
import com.xiaoyu.exception.InvalidParametersException;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * <p>
 * Excel表格工具类，处理常用的Excel表格操作<br/>
 * 如果使用上的bug和问题请联系！
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-13 14:04
 */
public class ExcelUtils {


    /**
     * 表格每个字符所占的宽度
     */
    private static final int CELL_WIDTH = 550;


    /**
     * 导出Excel
     *
     * @param sheetName 名称
     * @param title     注意! Map的Key是：实体的字段名称  value是：标题，用于创建表格第一行（如：编号，创建时间）
     * @param list      表格数据，请使用泛型！
     * @return HSSFWorkbook
     * @throws InvocationTargetException  调用目标异常
     * @throws IllegalAccessException     非法访问异常
     * @throws NoSuchMethodException      没有这样的方法异常
     * @throws InvalidParametersException 参数错误异常
     * @throws ParseException             转换异常
     */
    public static HSSFWorkbook exportExcel(String sheetName, LinkedHashMap<String, String> title, List<?> list)
            throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InvalidParametersException, ParseException {
        //校验参数
        String verificationStr = verification(sheetName, title, list);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb = createHSSFWorkbook(sheetName);
        HSSFCell cell;
        HSSFSheet sheet = wb.getSheet(sheetName);
        HSSFRow row = sheet.createRow(0);

        //设置表头
        int c = 0;
        for (String string : title.keySet()) {
            cell = row.createCell(c);
            cell.setCellValue(title.get(string));
            cell.setCellStyle(getCellStyle(wb));
            c++;
        }
        //Class对象
        Class<?> beanClass = list.get(0).getClass();
        //实体对象
        Object object;
        Method method;
        String field;
        Map<String, Map<String, String>> maps = stringMapMap(beanClass);
        Map<String, String[]> mapByFiledValue = analysis(list.get(0));

        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            //获取对象
            object = list.get(i);
            c = 0;
            for (String string : title.keySet()) {
                method = beanClass.getDeclaredMethod("get" + DataUtils.captureName(string));
                field = String.valueOf(method.invoke(object));
                if (maps.get(string) != null) {
                    //该字段有注解
                    String fieldInMap = maps.get(string).get(field);
                    field = DataUtils.isEmpty(fieldInMap) ? field : fieldInMap;

                    if (mapByFiledValue.get(string) != null) {
                        //字段包含@FiledValue注解
                        field = mapByFiledValue.get(string)[0] + field + mapByFiledValue.get(string)[1];
                    }
                    row.createCell(c).setCellValue(field);
                } else {
                    if (mapByFiledValue.get(string) != null) {
                        //字段包含@FiledValue注解
                        field = mapByFiledValue.get(string)[0] + field + mapByFiledValue.get(string)[1];
                    }
                    //判断返回值
                    if (method.getReturnType() == double.class) {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "0.0" : field);
                    } else if (method.getReturnType() == int.class || method.getReturnType() == Integer.class) {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "0" : field);
                    } else if (method.getReturnType() == Date.class) {
                        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        String dateString = formatter.format(formatter.parse(field));
                        row.createCell(c).setCellValue(DataUtils.isEmpty(dateString) ? "" : field);
                    } else if (method.getReturnType() == Timestamp.class) {
                        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        row.createCell(c).setCellValue(df.format(Timestamp.valueOf(field)));
                    } else if ("null".equals(field)) {
                        row.createCell(c).setCellValue("");
                    } else {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "" : field);
                    }
                }
                c++;
            }
        }
        //设定自动宽度
        int i = 0;
        for (String string : title.keySet()) {
            sheet.setColumnWidth(i, string.length() * CELL_WIDTH);
            i++;
        }
        return wb;
    }


    /**
     * 通过List导出Excel表格
     *
     * @param sheetName 文件名
     * @param lists     数据表格
     * @return HSSFWorkbook
     * @throws InvalidParametersException 参数错误异常
     */
    public static HSSFWorkbook exportExcel(String sheetName, List<List<String>> lists) throws InvalidParametersException {
        //校验参数
        String verificationStr = verification(sheetName, lists);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb = createHSSFWorkbook(sheetName);
        HSSFSheet sheet = wb.getSheet(sheetName);
        HSSFRow row = sheet.createRow(0);

        HSSFCell cell;
        for (int j = 0; j < lists.get(0).size(); j++) {
            cell = row.createCell(j);
            cell.setCellValue(lists.get(0).get(j));
            cell.setCellStyle(getCellStyle(wb));
        }
        for (int i = 1; i < lists.size(); i++) {
            row = sheet.createRow(i);
            //获取对象
            List<String> list = lists.get(i);
            for (int d = 0; d < list.size(); d++) {
                row.createCell(d).setCellValue(list.get(d));
            }
        }
        // 设定自动宽度
        for (int i = 0; i < lists.get(0).size(); i++) {
            sheet.setColumnWidth(i, lists.get(0).get(i).length() * CELL_WIDTH);
            sheet.autoSizeColumn(i);
        }
        return wb;
    }


    /**
     * 通过List Object[] 导出Excel表格
     *
     * @param sheetName  文件名
     * @param objectList 数据表格
     * @param titles     标题
     * @return HSSFWorkbook
     * @throws InvalidParametersException 参数错误异常
     */
    public static HSSFWorkbook exportExcel(String sheetName, List objectList, String[] titles) throws InvalidParametersException {
        //校验参数
        String verificationStr = verification(sheetName, objectList);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb = createHSSFWorkbook(sheetName);
        HSSFSheet sheet = wb.createSheet(sheetName);
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell;

        //标题
        for (int j = 0; j < titles.length; j++) {
            cell = row.createCell(j);
            cell.setCellValue(titles[j]);
            cell.setCellStyle(getCellStyle(wb));
        }
        for (int i = 1; i < objectList.size(); i++) {
            row = sheet.createRow(i);
            //获取对象
            Object[] cellObject = (Object[]) objectList.get(i);
            for (int d = 0; d < cellObject.length; d++) {
                row.createCell(d).setCellValue(cellObject[d].toString());
            }
        }
        // 设定自动宽度，等表格完善后再设定
        for (int i = 0; i < titles.length; i++) {
            sheet.setColumnWidth(i, titles[i].length() * CELL_WIDTH);
        }
        return wb;
    }

    /**
     * 解析注解
     *
     * @param clazz 类Class对象
     * @return map
     * @throws InvalidParametersException Exception
     */
    private static Map<String, Map<String, String>> stringMapMap(Class<?> clazz) throws InvalidParametersException {
        Field[] fields = clazz.getDeclaredFields();
        Map<String, Map<String, String>> maps = new HashMap<>(fields.length);
        //每次循环 + 2
        int region = 2;
        if (fields.length > 0) {
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelField.class)) {
                    ExcelField annotation = field.getAnnotation(ExcelField.class);
                    //参数数组
                    String[] param = annotation.value();
                    Map<String, String> map = new HashMap<>(param.length / 2);
                    if (param.length % 2 != 0) {
                        throw new InvalidParametersException("注解参数格式错误！请检查！");
                    }
                    for (int i = 0; i < param.length; i += region) {
                        map.put(param[i], param[i + 1]);
                    }
                    maps.put(field.getName(), map);
                }
            }
        }
        return maps;
    }

    /**
     * 解析 @FiledValue 注解
     *
     * @param object 类对象
     * @return Map<String, String>
     */
    private static Map<String, String[]> analysis(Object object) {
        Field[] fields = object.getClass().getDeclaredFields();
        Map<String, String[]> resMap = new HashMap<>(fields.length);
        if (fields.length > 0) {
            for (Field field : fields) {
                if (field.isAnnotationPresent(FiledValue.class)) {
                    FiledValue annotation = field.getAnnotation(FiledValue.class);
                    //参数数组
                    String beginAppend = annotation.beginAppend();
                    String endAppend = annotation.endAppend();
                    resMap.put(field.getName(), new String[]{beginAppend, endAppend});
                }
            }
        }
        return resMap;
    }


    /**
     * 读取Excel文件，将Excel文件中的数据封装到JAVA bean中，方便操作！<br/>
     * 使用须知：Excel中的 数据顺序需要去 实体类中 字段的顺序一致<br/>
     * Excel中的第一个对应实体字段中的第一个字段……
     *
     * @param o         需要承载的对象
     * @param file      导入的文件
     * @param uuid      实体类是否有UUID，是否实现的序列化
     * @param startRows 开始读取的行数，因为一般第一行是标题
     * @return List 实体集合
     */
    public static List<Object> importExcel(Object o, File file, boolean uuid, int startRows) throws Exception {
        return readExcel(o, file, uuid, startRows);
    }

    /**
     * 读取Excel文件，默认从第二行读取，将Excel文件中的数据封装到JAVA bean中，方便操作！<br/>
     * 使用须知：Excel中的 数据顺序需要去 实体类中 字段的顺序一致<br/>
     * Excel中的第一个对应实体字段中的第一个字段……
     *
     * @param o    需要承载的对象
     * @param file 导入的文件
     * @param uuid 实体类是否有UUID，是否实现的序列化
     * @return List 实体集合
     */
    public static List<Object> importExcel(Object o, File file, boolean uuid) throws Exception {
        return readExcel(o, file, uuid, 2);
    }

    /**
     * 读取Excel文件，数据封装到List
     *
     * @param file      导入的文件
     * @param startRows 开始读取的行数，因为一般第一行是标题
     * @return List<List < String>>  List集合
     */
    public static List<List<String>> importExcel(File file, int startRows) throws IOException, InvalidParametersException {
        return readExcel(file, startRows);
    }

    /**
     * 读取Excel文件，数据封装到List，默认从表格的第二行开始读取
     *
     * @param file 导入的文件
     * @return List<List < String>>  List集合
     */
    public static List<List<String>> importExcel(File file) throws IOException, InvalidParametersException {
        return readExcel(file, 2);
    }

    /**
     * 读取Excel文件，数据封装到List ojc[]，此方法可以选择 sheet
     *
     * @param file        读取文件
     * @param startRows   开始读取行数
     * @param sheetNumber sheet 工作表索引
     * @return List<Object [ ]>
     * @throws IOException                io
     * @throws InvalidParametersException 参数校验失败
     */
    public static List<Object[]> importExcel(File file, int startRows, int sheetNumber) throws IOException, InvalidParametersException {
        return readExcel(file, startRows, sheetNumber);
    }

    private static List<Object> readExcel(Object o, File file, boolean uuid, int startRows) throws Exception {
        //校验参数
        String verificationStr = verification(o, file, startRows);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb;
        wb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = wb.getSheetAt(0);
        List<Object> list = new ArrayList<>(sheet.getLastRowNum());
        Method method;
        Class<?> cc = o.getClass();

        Map<String, Map<String, String>> maps = stringMapMap(cc);
        Field[] fields;
        if (uuid) {
            fields = arraySubtract(cc.getDeclaredFields());
        } else {
            fields = cc.getDeclaredFields();
        }
        for (int j = (startRows - 1); j < sheet.getLastRowNum() + 1; j++) {
            //创建实列
            Object newInstance = cc.newInstance();
            HSSFRow row = sheet.getRow(j);
            if (row.getLastCellNum() != fields.length) {
                throw new Exception("fields 的长度与Excel 表格行长度不匹配，在Excel：" + j + "行");
            }
            for (int i = 0; i < row.getLastCellNum(); i++) {
                //字段名称
                String fieldName = fields[i].getName();
                if ("serialVersionUID".equals(fieldName)) {
                    break;
                }
                HSSFCell cell = row.getCell(i);
                //属性类型
                Class<?> fieldType = fields[i].getType();
                //格数据
                String cellValue;
                try {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                } catch (Exception e) {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }

                if (maps.get(fieldName) != null) {
                    //字段有注解 ，判断是否符合注解中的条件
                    String fieldInMap = maps.get(fieldName).get(cellValue);
                    cellValue = DataUtils.isEmpty(fieldInMap) ? cellValue : fieldInMap;
                }

                method = cc.getDeclaredMethod("set" + DataUtils.captureName(fieldName), fieldType);
                if (fieldType == int.class) {
                    method.invoke(newInstance, Integer.parseInt(cellValue));
                } else if (fieldType == Long.class) {
                    method.invoke(newInstance, Long.parseLong(cellValue));
                } else if (fieldType == double.class) {
                    method.invoke(newInstance, Double.parseDouble(cellValue));
                } else if (fieldType == Float.class) {
                    method.invoke(newInstance, Float.parseFloat(cellValue));
                } else if (fieldType == Date.class) {
                    try {
                        method.invoke(newInstance, new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(cellValue));
                    } catch (Exception e) {
                        method.invoke(newInstance, new SimpleDateFormat("yyyy-MM-dd").parse(cellValue));
                    }
                } else if (fieldType == byte.class) {
                    method.invoke(newInstance, Byte.parseByte(cellValue));
                } else if (fieldType == short.class) {
                    method.invoke(newInstance, Short.parseShort(cellValue));
                } else if (method.getReturnType() == Timestamp.class) {
                    method.invoke(newInstance, Timestamp.valueOf(cellValue));
                } else {
                    method.invoke(newInstance, cellValue);
                }
                //如有缺少的类型请自行补上
            }
            list.add(newInstance);
        }
        return list;
    }

    private static List<List<String>> readExcel(File file, int startRows) throws IOException, InvalidParametersException {
        //校验参数
        String verificationStr = verification(file, startRows);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb;
        wb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = wb.getSheetAt(0);
        List<List<String>> lists = new ArrayList<>(sheet.getLastRowNum());
        for (int j = (startRows - 1); j < sheet.getLastRowNum() + 1; j++) {
            HSSFRow row = sheet.getRow(j);
            List<String> list = new ArrayList<>(row.getLastCellNum());
            for (int i = 0; i < row.getLastCellNum(); i++) {
                HSSFCell cell = row.getCell(i);
                //格数据
                String cellValue;
                try {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                } catch (Exception e) {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                list.add(cellValue);
            }
            lists.add(list);
        }
        return lists;
    }

    private static List<Object[]> readExcel(File file, int startRows, int sheetNumber) throws IOException, InvalidParametersException {
        //校验参数
        String verificationStr = verification(file, startRows);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        if (sheetNumber < 0) {
            sheetNumber = 1;
        }
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = wb.getSheetAt(sheetNumber);
        List<Object[]> objects = new ArrayList<>(sheet.getLastRowNum());
        for (int j = (startRows - 1); j < sheet.getLastRowNum() + 1; j++) {
            HSSFRow row = sheet.getRow(j);
            Object[] object = new Object[row.getLastCellNum()];
            for (int i = 0; i < row.getLastCellNum(); i++) {
                HSSFCell cell = row.getCell(i);
                //格数据
                String cellValue;
                try {
                    cellValue = String.valueOf(cell.getRichStringCellValue());
                } catch (Exception e) {
                    cellValue = String.valueOf(cell.getNumericCellValue());
                }
                object[i] = cellValue;
            }
            objects.add(object);
        }
        return objects;
    }

    /**
     * 移除数组第一位
     *
     * @param fields 数组
     * @return 新数组
     */
    private static Field[] arraySubtract(Field[] fields) {
        Field[] fields1 = new Field[fields.length - 1];
        System.arraycopy(fields, 1, fields1, 0, fields1.length);
        return fields1;
    }

    /**
     * 验证
     *
     * @param o         实体参数
     * @param file      文件
     * @param startRows 起始行数
     * @return String
     */
    private static String verification(Object o, File file, int startRows) {
        if (o == null) {
            return "实体类参数不能为空 ！";
        }
        if (file == null) {
            return "导入文件不能为空 ！";
        }
        if (startRows < 1) {
            return "读取起始行数不能小于1 ！";
        }
        return null;
    }

    /**
     * 验证
     *
     * @param file      文件
     * @param startRows 起始行数
     * @return boolean
     */
    private static String verification(File file, int startRows) {
        if (file == null) {
            return "导入文件不能为空 ！";
        }
        if (startRows < 1) {
            return "读取起始行数不能小于1 ！";
        }
        return null;
    }

    /**
     * 验证格式是否正确
     *
     * @param sheetName sheetName
     * @param title     title
     * @param list      list
     * @return boolean
     */
    private static String verification(String sheetName, Map<String, String> title, List<?> list) {
        if (DataUtils.isEmpty(sheetName)) {
            return "导出失败！原因：sheetName 不能为空！";
        }
        if (title == null || title.isEmpty()) {
            return "导出失败！原因：Map不能为空！";
        }
        if (DataUtils.isEmpty(list)) {
            return "导出失败！原因：List 表格数据不能为空！";
        }
        return null;
    }

    private static String verification(String sheetName, List<?> lists) {
        if (DataUtils.isEmpty(sheetName)) {
            return "导出失败！原因：sheetName 不能为空！";
        }
        if (DataUtils.isEmpty(lists)) {
            return "导出失败！原因：List 表格数据不能为空！";
        }
        return null;
    }

    private static HSSFWorkbook createHSSFWorkbook(String sheetName) {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);
        //固定标题栏
        sheet.createFreezePane(0, 1, 0, 1);
        return wb;
    }

    private static HSSFCellStyle getCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

}
