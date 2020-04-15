package com.xiaoyu.utils;

import java.util.List;

public class DataUtils {

	public static boolean isEmpty(String string) {
        return string == null || "".equals(string);
    }

	/**
	 * 将字符串首字母大写
	 *
	 * @param name 需要处理的字符串
	 * @return 处理后的字符串
	 */
	public static String captureName(String name) {
		char[] cs = name.toCharArray();
		cs[0] -= 32;
		return String.valueOf(cs);
	}

	/**
	 * 校验集合是否是null 或者长度为0
	 *
	 * @param list  list
	 * @return boolean
	 */
	public static boolean isEmpty(List list) {
		return list == null || list.size() == 0;
	}

}
