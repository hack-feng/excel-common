package com.common.excel.util.excel.bean;

import lombok.Data;

import java.io.Serializable;

/**
 * @author zhangfuzeng
 * @date 2021/12/10
 */
@Data
public class UserBean implements Serializable {


    private String className;

    private String userName;

    private String sex;

    private Integer age;

    private Integer scope;

    private String remark;

    private String remark2;
}
