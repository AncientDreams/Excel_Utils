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
     */
    public static HSSFWorkbook exportExcel(String sheetName, LinkedHashMap<String, String> title, List<?> list)
            throws InvocationTargetException, IllegalAccessException, NoSuchMethodException, InvalidParametersException {
        //校验参数
        String verificationStr = verification(sheetName, title, list);
        if (verificationStr != null) {
            throw new InvalidParametersException(verificationStr);
        }
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);
        //固定标题栏
        sheet.createFreezePane(0, 1, 0, 1);
        HSSFRow row = sheet.createRow(0);
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFCell cell;
//        log.info(LOG_STR + "正在导出Excel表格，sheetName：{}，标题长度：{}，表格数据条数：{}", sheetName, title.size(), list.size());
        int c = 0;
        for (String string : title.keySet()) {
            cell = row.createCell(c);
            cell.setCellValue(title.get(string));
            cell.setCellStyle(style);
            c++;
        }
        //获取集合中的第一个属性Class对象，其他属性Class属性都一致
        Class<?> beanClass = list.get(0).getClass();
        //实体对象
        Object object;
        Method method;
        //值
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

                    if(mapByFiledValue.get(string) != null){
                        //字段包含@FiledValue注解
                        field = mapByFiledValue.get(string)[0] + field + mapByFiledValue.get(string)[1];
                    }
                    row.createCell(c).setCellValue(field);
                } else {
                    if(mapByFiledValue.get(string) != null){
                        //字段包含@FiledValue注解
                        field = mapByFiledValue.get(string)[0] + field + mapByFiledValue.get(string)[1];
                    }
                    //判断返回值
                    if (method.getReturnType() == double.class) {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "0.0" : field);
                    } else if (method.getReturnType() == int.class || method.getReturnType() == Integer.class) {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "0" : field);
                    } else if ("null".equals(field)) {
                        row.createCell(c).setCellValue("");
                    } else {
                        row.createCell(c).setCellValue(DataUtils.isEmpty(field) ? "" : field);
                    }
                }
                c++;
            }
        }
        // 设定自动宽度
        for (int i = 0; i < title.size(); i++) {
            sheet.autoSizeColumn(i);
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
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet(sheetName);
        //固定标题栏
        sheet.createFreezePane(0, 1, 0, 1);
        HSSFRow row = sheet.createRow(0);
        HSSFCellStyle style = wb.createCellStyle();
        HSSFFont font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFCell cell;
        for (int j = 0; j < lists.get(0).size(); j++) {
            cell = row.createCell(j);
            cell.setCellValue(lists.get(0).get(j));
            cell.setCellStyle(style);
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
            sheet.autoSizeColumn(i);
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
                    for (int i = 0; i < param.length; i += 2) {
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
     * @param object  类对象
     * @return Map<String,String>
     */
    private static Map<String,String[]> analysis(Object object) {
        Field[] fields =  object.getClass().getDeclaredFields();
        Map<String,String[]> resMap = new HashMap<>(fields.length);
        if (fields.length > 0) {
            for (Field field : fields) {
                if (field.isAnnotationPresent(FiledValue.class)) {
                    FiledValue annotation = field.getAnnotation(FiledValue.class);
                    //参数数组
                    String beginAppend = annotation.beginAppend();
                    String endAppend = annotation.endAppend();
                    resMap.put(field.getName(), new String[]{beginAppend,endAppend});
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
                String cellValue = String.valueOf(cell.getRichStringCellValue());

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
                String cellValue = String.valueOf(cell.getRichStringCellValue());
                list.add(cellValue);
            }
            lists.add(list);
        }
        return lists;
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

    private static String verification(String sheetName, List<List<String>> lists) {
        if (DataUtils.isEmpty(sheetName)) {
            return "导出失败！原因：sheetName 不能为空！";
        }
        if (DataUtils.isEmpty(lists)) {
            return "导出失败！原因：List 表格数据不能为空！";
        }
        return null;
    }

}
