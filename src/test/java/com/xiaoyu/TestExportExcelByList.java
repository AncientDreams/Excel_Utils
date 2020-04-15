package com.xiaoyu;

import com.xiaoyu.entity.Students;
import com.xiaoyu.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * <p>
 * 测试导出通过 List
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-15 16:38
 */
public class TestExportExcelByList {

    @Test
    public void exportExcel() throws IOException {
        List<List<String>> lists = new ArrayList<>();
        List<String> list = new ArrayList<>();
        list.add("姓名");list.add("年龄");list.add("姓别");
        List<String> list1= new ArrayList<>();
        list1.add("姓名1");list1.add("年龄1");list1.add("姓别1");
        List<String> list2= new ArrayList<>();
        list2.add("姓名2");list2.add("年龄2");list2.add("姓别2");
        List<String> list3= new ArrayList<>();
        list3.add("姓名3");list3.add("年龄3");list3.add("姓别3");
        lists.add(list); lists.add(list1); lists.add(list2); lists.add(list3);
        File f = new File("exportExcelByList1.xls");
        OutputStream out =new FileOutputStream(f);
        HSSFWorkbook sheets = ExcelUtils.exportExcel("测试导出List", lists);
        sheets.write(out);
        out.flush();
        out.close();
    }

}
