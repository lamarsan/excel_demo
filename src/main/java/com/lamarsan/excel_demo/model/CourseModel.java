package com.lamarsan.excel_demo.model;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import cn.afterturn.easypoi.excel.annotation.ExcelEntity;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.util.List;

/**
 * className: CourseModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 17:53
 */
@Data
@ExcelTarget("courseEntity")
@NoArgsConstructor
@AllArgsConstructor
public class CourseModel implements Serializable {
    /**
     * 主键
     */
    private String id;
    /**
     * 课程名称
     */
    @Excel(name = "课程名称", orderNum = "1",needMerge = true, width = 25)
    private String name;
    /**
     * 老师主键
     */
    @ExcelEntity(id = "absent")
    private TeacherModel mathTeacher;

    @ExcelCollection(name = "学生", orderNum = "4")
    private List<StudentModel> students;

}
