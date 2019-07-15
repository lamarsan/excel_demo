package com.lamarsan.excel_demo.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;

/**
 * className: TeacherModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 17:55
 */
@Data
@ExcelTarget("teacherEntity")
@AllArgsConstructor
@NoArgsConstructor
public class TeacherModel implements Serializable {
    private String id;
    /** name */
    @Excel(name = "主讲老师_major,代课老师_absent",needMerge = true, orderNum = "1", isImportField = "true_major,true_absent")
    private String name;
}
