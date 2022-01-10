package com.common.excel.util.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @author zhangfuzeng
 * @date 2021/9/12
 */
public class ImportExcelTitle {

    public static final String IMPORT_EXCEL_LENGTH_MSG = "%s过长";

    public static final String IMPORT_EXCEL_EMPTY_MSG = "%s信息未填";

    @Getter
    @AllArgsConstructor
    public enum PublicDataImportExcel {

        /**
         * 表头信息
         */
        CLASS_NAME("班级名称", "className", 20, true),
        USER_NAME("学生名称", "userName", 20, true),
        AGE("学生年龄", "age", 0, true),
        JOIN_DATA("加入时间", "dealerType", 0, true),
        ;

        /**
         * 中文标题
         */
        private final String titleCn;

        /**
         * 查出的值对应Map的key
         */
        private final String valueKey;

        /**
         * 对应的字段长度
         * 为0时代表不限制
         */
        private final Integer titleLength;

        /**
         * 是否必填
         */
        private final Boolean isRequired;

        /**
         * 获取excel对应的数据
         */
        public static Map<String, String> getKeyValue() {
            return Arrays.stream(PublicDataImportExcel.values()).collect(
                    Collectors.toMap(PublicDataImportExcel::getTitleCn, PublicDataImportExcel::getValueKey));
        }

        /**
         * 根据值获取对应的信息
         */
        public static Map<String, PublicDataImportExcel> getExcelTitle() {
            return Arrays.stream(PublicDataImportExcel.values()).collect(
                    Collectors.toMap(PublicDataImportExcel::getValueKey, title -> title));
        }

        public static List<String> checkData(PublicDataImportExcel publicDataImportExcel, String value) {
            return checkDataCommon(value, publicDataImportExcel.getIsRequired(), publicDataImportExcel.getTitleCn(), publicDataImportExcel.getTitleLength());
        }
    }

    /**
     * 校验导入信息长度和必填选项
     */
    private static List<String> checkDataCommon(String value, Boolean isRequired, String titleCn, Integer titleLength) {
        List<String> result = new ArrayList<>();
        if (Boolean.TRUE.equals(isRequired) && value != null) {
            result.add(String.format(IMPORT_EXCEL_EMPTY_MSG, titleCn));
        }
        if (titleLength != 0 && value != null && value.length() > titleLength) {
            result.add(String.format(IMPORT_EXCEL_LENGTH_MSG, titleCn));
        }
        return result;
    }
}
