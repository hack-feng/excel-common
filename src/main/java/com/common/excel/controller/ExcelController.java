package com.common.excel.controller;

import com.alibaba.fastjson.JSON;
import com.common.excel.util.excel.ExportExcelTitle;
import com.common.excel.util.excel.ExportExcelUtil;
import com.common.excel.util.excel.ImportExcelTitle;
import com.common.excel.util.excel.ImportExcelUtil;
import com.common.excel.util.excel.bean.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * @author zhangfuzeng
 * @date 2021/12/13
 */
@Slf4j
@RestController
@RequestMapping(value = "/excel")
public class ExcelController {

    /**
     * 测试导出excel的格式
     */
    @GetMapping("/testExcel")
    public void testExcel(HttpServletRequest request, HttpServletResponse response) {
        try {
            ExportExcelUtil.updateNameUnicode(request, response, "测试");
            OutputStream out = response.getOutputStream();
            ExportExcelUtil.testExcel(out);
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 单个sheet单个table导出
     */
    @GetMapping("/exportUser")
    public void exportUser(HttpServletRequest request, HttpServletResponse response) {

        List<GradeBean> list = getGrade();
        String excelName = "用户信息";
        try {
            ExportExcelUtil.updateNameUnicode(request, response, excelName);
            ExportTableBean exportBean = ExportExcelTitle.GradeExcel.getValue(list, excelName, ExportExcelTheme.ORANGE);
            OutputStream out = response.getOutputStream();
            ExportExcelUtil.exportExcel(exportBean, "测试sheet页", "123456", out);
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 多个sheet，多个table导出
     */
    @GetMapping("/exportMoreUser")
    public void exportMoreUser(HttpServletRequest request, HttpServletResponse response) {

        List<ExportSheetBean> sheetBeans = new ArrayList<>();
        String excelName = "用户信息";
        try {
            ExportExcelUtil.updateNameUnicode(request, response, excelName);

            // 第1个sheet页
            ExportSheetBean sheetBean = new ExportSheetBean();
            sheetBean.setSheetName("学生信息");
            sheetBean.setProtectSheet("123456");
            List<ExportTableBean> exportTableBeans = new ArrayList<>();
            // 第1个sheet页, 第1个table
            exportTableBeans.add(ExportExcelTitle.UserExcel.getValue(getUserList(), excelName, ExportExcelTheme.BLUE));
            // 第1个sheet页, 第2个table
            exportTableBeans.add(ExportExcelTitle.UserExcel.getValue(getUserList(), excelName, ExportExcelTheme.ORANGE));
            sheetBean.setList(exportTableBeans);

            // 第2个sheet页
            ExportSheetBean sheetBean2 = new ExportSheetBean();
            sheetBean2.setSheetName("学生成绩");
            List<ExportTableBean> exportTableBeans2 = new ArrayList<>();
            // 第2个sheet页, 第1个table
            exportTableBeans2.add(ExportExcelTitle.GradeExcel.getValue(getGrade(), excelName, ExportExcelTheme.BLUE));
            // 第2个sheet页, 第2个table
            exportTableBeans2.add(ExportExcelTitle.GradeExcel.getValue(getGrade(), excelName, ExportExcelTheme.GREEN));
            sheetBean2.setList(exportTableBeans2);

            sheetBeans.add(sheetBean);
            sheetBeans.add(sheetBean2);

            OutputStream out = response.getOutputStream();
            ExportExcelUtil.exportExcelMoreSheetMoreTable(sheetBeans, out);
            out.flush();
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/importUser")
    public void importUser(@RequestParam(value = "file") MultipartFile fileToUpload) {
        List<Map<String, String>> excelList;
        try {
            excelList = ImportExcelUtil.parseExcel(
                    fileToUpload.getInputStream(),
                    fileToUpload.getOriginalFilename(),
                    ImportExcelTitle.PublicDataImportExcel.getKeyValue());
        } catch (IOException e) {
            log.error("解析excel时失败" + e.getMessage());
            throw new RuntimeException("excel解析失败");
        }
        log.info(JSON.toJSONString(excelList));
    }

    /**
     * 用户信息
     */
    private List<UserBean> getUserList() {
        List<UserBean> userBeans = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            for (int j = 0; j < 5; j++) {
                try {
                    Thread.sleep(10);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                userBeans.add(getUser(i + "班"));
            }
        }
        return userBeans;
    }

    private UserBean getUser(String className) {
        String[] userName = {"大娃", "二娃", "三娃", "四娃", "五娃", "六娃", "七娃"};
        String[] remark = {"打球", "跑步", "跳舞", "唱歌", ""};
        Random r = new Random();
        UserBean userBean = new UserBean();
        userBean.setClassName(className);
        userBean.setUserName(userName[r.nextInt(7)]);
        userBean.setSex(r.nextInt(2) == 1 ? "女" : "男");
        userBean.setAge(r.nextInt(5) + 10);
        userBean.setScope(r.nextInt(10));
        userBean.setRemark(remark[r.nextInt(5)]);
        userBean.setRemark2(remark[r.nextInt(5)]);
        return userBean;
    }

    private List<GradeBean> getGrade() {
        List<GradeBean> gradeBeans = new ArrayList<>();
        for (int i = 1; i < 4; i++) {
            for (int j = 0; j < 5; j++) {
                try {
                    Thread.sleep(10);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                gradeBeans.add(getGrade(i, j));
            }
        }
        return gradeBeans;
    }

    private GradeBean getGrade(Integer indexClass, Integer indexName) {
        String[] userName = {"大娃", "二娃", "三娃", "四娃", "五娃"};
        Random r = new Random();
        GradeBean gradeBean = new GradeBean();
        gradeBean.setClassName(indexClass + "班");
        gradeBean.setUserName(userName[indexName]);
        gradeBean.setMathGrade(r.nextInt(30) + 70);
        gradeBean.setEnglishGrade(r.nextInt(40) + 60);
        gradeBean.setChinaGrade(r.nextInt(20) + 80);
        return gradeBean;
    }
}
