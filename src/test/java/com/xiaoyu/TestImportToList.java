package com.xiaoyu;

import com.xiaoyu.exception.InvalidParametersException;
import com.xiaoyu.utils.ExcelUtils;
import org.junit.Test;
import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * <p>
 * 测试将数据导入到 List 中
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-15 17:09
 */
public class TestImportToList {

    @Test
    public void testImportToBean() throws IOException, InvalidParametersException {
        File file  = new File("exportExcelByList2.xls");
        //导入一样可以注解转换，这里不做示范,默认不读取第一行。
        List list =  ExcelUtils.importExcel(file);
        System.out.println(list.get(0).toString());
        System.out.println(list.size());
    }
}
