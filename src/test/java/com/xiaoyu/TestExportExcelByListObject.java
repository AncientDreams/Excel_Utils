package com.xiaoyu;

import com.xiaoyu.exception.InvalidParametersException;
import com.xiaoyu.utils.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * <p>
 * 测试通过 List<Object> 导出Excel报表
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-06-19 15:07
 */
public class TestExportExcelByListObject {

    @Test
    public void exportExcel() throws IOException, InvalidParametersException {
        List<Object[]> list = new ArrayList<>();
        String[] titles = new String[]{"姓名", "性别", "年龄"};

        Object[] o2 = new Object[3];
        o2[0] = "张三";
        o2[1] = "男";
        o2[2] = "20";

        Object[] o3 = new Object[3];
        o3[0] = "小红";
        o3[1] = "女";
        o3[2] = "20";
        list.add(o2);
        list.add(o3);
        File f = new File("exportExcelByList3.xls");
        OutputStream out = new FileOutputStream(f);
        HSSFWorkbook sheets = ExcelUtils.exportExcel("测试", list, titles);
        sheets.write(out);
        out.flush();
        out.close();
    }
}
