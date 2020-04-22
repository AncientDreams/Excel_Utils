package com.xiaoyu.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <p>
 * 该注解配合ExcelUtils使用 {@link  com.xiaoyu.utils.ExcelUtils} <br/>
 * 作用：在 [导出] 报表对单个属性值进行操作，如：在开始或者结束时添加字符等。  <br/>
 * 使用方式：在需要转换的实体类的字段属性上加上该注解！<br/>
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-22 11:50
 */
@Retention(value = RetentionPolicy.RUNTIME)
@Target(value = ElementType.FIELD)
public @interface FiledValue {

    /**
     * 在开始处追加字符
     */
    String beginAppend() default "";

    /**
     * 在结束处追加字符
     */
    String endAppend() default "";
}
