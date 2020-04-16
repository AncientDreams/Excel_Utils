package com.xiaoyu.exception;

/**
 * <p>
 *  参数无效抛出此异常
 * </p>
 *
 * @author ZhangXianYu   Email: 1600501744@qq.com
 * @since 2020-04-16 15:31
 */
public class InvalidParametersException extends Exception {

    public InvalidParametersException(String message) {
        super(message);
    }
}
