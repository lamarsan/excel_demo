package com.lamarsan.excel_demo.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.util.Date;

/**
 * className: ExcelImportModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 16:11
 */
@Data
public class ExcelImportModel {
    @Excel(name = "id")
    private String id;
    //可以自动替换excel中内容
    @Excel(name = "性别", replace = {"男_man", "女_woman"})
    private String sex;
    @Excel(name = "年龄")
    private String age;
    @Excel(name = "姓名")
    private String name;

}
