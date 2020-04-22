package com.xiaoyu;

import com.xiaoyu.entity.Student;
import com.xiaoyu.exception.InvalidParametersException;
import com.xiaoyu.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * <p>
 * 测试导出通过 Java Bean
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-15 17:03
 */
public class TestExportExcelByBean {

    @Test
    public void exportExcel1() throws IOException, InvalidParametersException, NoSuchMethodException,
            IllegalAccessException, InvocationTargetException {
        LinkedHashMap<String,String> map = new LinkedHashMap<>();
        map.put("name", "姓名");
        map.put("sex", "姓别");
        map.put("age", "年龄");

        List list = new ArrayList();
        //虽然填写的是数字，但是下载后会根据注解转换
        Student student  = new Student("张","00",22);
        Student student1  = new Student("张3","01",22);
        list.add(student); list.add(student1);
        File f = new File("exportExcelByList2.xls");
        OutputStream out = new FileOutputStream(f);
        HSSFWorkbook sheets = ExcelUtils.exportExcel("s",map,list);
        sheets.write(out);
        out.flush();
        out.close();
    }
}
