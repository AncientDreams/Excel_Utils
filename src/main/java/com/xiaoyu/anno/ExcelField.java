package com.xiaoyu.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <p>
 * 该注解配合ExcelUtils使用 {@link  com.xiaoyu.utils.ExcelUtils} <br/>
 * 作用：在导入、导出 报表的时候，做转换。如： 将 00转换为 成功！  <br/>
 * 使用方式：在需要转换的实体类的字段属性上加上该注解，并且value的长度需要2的倍数<br/>
 * 如： @ExcelField(value = {"00","成功","01","失败"})
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-14 9:37
 */
@Retention(value = RetentionPolicy.RUNTIME)
@Target(value = ElementType.FIELD)
public @interface ExcelField {

    String[] value();
}
