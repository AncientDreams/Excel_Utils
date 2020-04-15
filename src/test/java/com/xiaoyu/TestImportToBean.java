package com.xiaoyu;

import com.xiaoyu.entity.Students;
import com.xiaoyu.utils.ExcelUtils;
import org.junit.Test;

import java.io.File;
import java.util.List;

/**
 * <p>
 * 测试讲数据导入到Java对象中
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-15 17:05
 */
public class TestImportToBean {

    @Test
    public void testImportToBean(){
        File file  = new File("exportExcelByList2.xls");
        Students students = new Students();
        //导入 ，导入一样可以注解转换，这里 不做示范
        List list =  ExcelUtils.importExcel(students, file, false);
        System.out.println(list.get(0).toString());
        System.out.println(list.size());
    }

}
