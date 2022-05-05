package com.yyz.util.constant;

/**
 * 项目常量类
 * @Author: yyz
 * @Date: 2022/4/24 15:54
 */
public class Constants {

    /**
     * 占位符第一个字符
     */
    public static final String PREFIX_FIRST = "$";

    /**
     * 占位符第二个字符
     */
    public static final String PREFIX_SECOND = "{";

    /**
     * 占位符的前缀
     */
    public static final String PLACEHOLDER_PREFIX = PREFIX_FIRST + PREFIX_SECOND;

    /**
     * 表格中需要动态添加行的独特标记
     */
    public static final String ADD_ROW_TEXT = "tbAddRow:";

    /**
     * 表格中占位符的开头 ${tbAddRow:  例如${tbAddRow:tb1}
     */
    public static final String ADD_ROW_FLAG = PLACEHOLDER_PREFIX + ADD_ROW_TEXT;

    /**
     * 占位符的后缀
     */
    public static final String PLACEHOLDER_SUFFIX = "}";

    /**
     * 图片占位符的前缀
     */
    public static final String PICTURE_PREFIX = PLACEHOLDER_PREFIX + "image:";
}
