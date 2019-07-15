package com.lamarsan.excel_demo.controller;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.lamarsan.excel_demo.common.ExportView;
import com.lamarsan.excel_demo.model.CourseModel;
import com.lamarsan.excel_demo.model.StudentModel;
import com.lamarsan.excel_demo.model.StudentReadModel;
import com.lamarsan.excel_demo.model.TeacherModel;
import com.lamarsan.excel_demo.utils.ExcelUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static cn.afterturn.easypoi.excel.entity.enmus.ExcelType.XSSF;

/**
 * className: ExcelDemoController
 * description: EasyPoi的使用
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 16:13
 */
@RestController
@RequestMapping("/excelReader")
public class EasyPoiController {
    /**
     * excel导入
     *
     * @param file 导入文件
     * @return excel对象数组
     */
    @PostMapping(value = "/excelImport")
    public Object importExcel(@RequestParam("file") MultipartFile file) {

        //接收导入数组
        List<StudentReadModel> studentModels = null;
        try {
            studentModels = ExcelImportUtil.importExcel(file.getInputStream(), StudentReadModel.class, new ImportParams());
        } catch (Exception e) {
            e.printStackTrace();
        }

        return studentModels;
    }

    /**
     * excel 下载
     *
     * @param response response只能使用一次
     */
    @RequestMapping(value = "/excelExport")
    public void exportExcel(HttpServletResponse response) {
        //mock 下载导出测试数据
        //ExcelExportModel export1 = new ExcelExportModel("5", "women", "18","李慧");
        //ExcelExportModel export2 = new ExcelExportModel("6", "men","20" ,"小城");
        //List<ExcelExportModel> list = new ArrayList<>();
        //list.add(export1);
        //list.add(export2);
        StudentModel studentModel = new StudentModel("1", "李慧", 1, new Date(), new Date());
        StudentModel studentMode2 = new StudentModel("2", "王慧", 1, new Date(), new Date());
        StudentModel studentMode3 = new StudentModel("3", "傅慧", 1, new Date(), new Date());
        StudentModel studentMode4 = new StudentModel("4", "周慧", 1, new Date(), new Date());
        StudentModel studentMode5 = new StudentModel("5", "程慧", 1, new Date(), new Date());
        List<StudentModel> list = new ArrayList<>();
        list.add(studentModel);
        list.add(studentMode2);
        List<StudentModel> list2 = new ArrayList<>();
        list2.add(studentMode3);
        list2.add(studentMode4);
        List<StudentModel> list3 = new ArrayList<>();
        list3.add(studentMode5);

        TeacherModel teacherModel = new TeacherModel("1", "青雉");
        TeacherModel teacherModel2 = new TeacherModel("2", "赤犬");
        TeacherModel teacherModel3 = new TeacherModel("3", "ppap");
        CourseModel courseModel = new CourseModel("1", "如何游泳", teacherModel, list);
        CourseModel courseModel2 = new CourseModel("2", "如何学习", teacherModel2, list2);
        CourseModel courseModel3 = new CourseModel("1", "如何rap", teacherModel3, list3);
        List<CourseModel> courseModels = new ArrayList<>();
        courseModels.add(courseModel);
        courseModels.add(courseModel2);
        courseModels.add(courseModel3);
        //参数配置
        //ExportParams params = new ExportParams();
        //此处设置ExcelType HSSF为excel2003版本，XSSF为excel2007版本
        //params.setType(ExcelType.XSSF);
        //Workbook workbook = ExcelExportUtil.exportExcel(params, ExcelExportModel.class, list);
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("2412312", "测试"),
                CourseModel.class, courseModels);
        ExcelUtil.downloadExcel(response, workbook, "计算机二班学生选课情况");
    }

    /**
     * excel 下载，合并单元格
     *
     * @param response response只能使用一次
     */
    @RequestMapping(value = "/exportMultipleExcel")
    public void exportMultipleExcel(HttpServletResponse response) {
        StudentModel studentModel = new StudentModel("1", "李慧", 1, new Date(), new Date());
        StudentModel studentMode2 = new StudentModel("2", "王慧", 1, new Date(), new Date());
        StudentModel studentMode3 = new StudentModel("3", "傅慧", 1, new Date(), new Date());
        StudentModel studentMode4 = new StudentModel("4", "周慧", 1, new Date(), new Date());
        StudentModel studentMode5 = new StudentModel("5", "程慧", 1, new Date(), new Date());
        List<StudentModel> list = new ArrayList<>();
        list.add(studentModel);
        list.add(studentMode2);
        List<StudentModel> list2 = new ArrayList<>();
        list2.add(studentMode3);
        list2.add(studentMode4);
        List<StudentModel> list3 = new ArrayList<>();
        list3.add(studentMode5);
        //学生的信息
        List<StudentModel> studentModelList = new ArrayList<>();
        studentModelList.add(studentModel);
        studentModelList.add(studentMode2);
        studentModelList.add(studentMode3);
        studentModelList.add(studentMode4);
        studentModelList.add(studentMode5);
        //课程的信息
        TeacherModel teacherModel = new TeacherModel("1", "青雉");
        TeacherModel teacherModel2 = new TeacherModel("2", "赤犬");
        TeacherModel teacherModel3 = new TeacherModel("3", "ppap");
        CourseModel courseModel = new CourseModel("1", "如何游泳", teacherModel, list);
        CourseModel courseModel2 = new CourseModel("2", "如何学习", teacherModel2, list2);
        CourseModel courseModel3 = new CourseModel("1", "如何rap", teacherModel3, list3);
        List<CourseModel> courseModelList = new ArrayList<>();
        courseModelList.add(courseModel);
        courseModelList.add(courseModel2);
        courseModelList.add(courseModel3);
        //教师的信息
        List<TeacherModel> teacherModelList = new ArrayList<>();
        teacherModelList.add(teacherModel);
        teacherModelList.add(teacherModel2);
        teacherModelList.add(teacherModel3);
        //参数配置 HSSF
        //ExportParams studentExportParams = new ExportParams();
        //studentExportParams.setSheetName("学生表");
        //XSSF
        List<Map<String, Object>> exportParamList = Lists.newArrayList();
        ExportView studentView = new ExportView(new ExportParams("学生表","表1",XSSF), studentModelList, StudentModel.class);
        ExportView courseView = new ExportView(new ExportParams("课程表","表2",XSSF), courseModelList, CourseModel.class);
        //ExportView teacherView = new ExportView(new ExportParams("教师表","表3",XSSF), teacherModelList, TeacherModel.class);
        List<ExportView> exportViews = new ArrayList<>();
        exportViews.add(studentView);
        //exportViews.add(teacherView);
        exportViews.add(courseView);
        for (ExportView view : exportViews) {
            Map<String, Object> valueMap = Maps.newHashMap();
            valueMap.put("title", view.getExportParams());
            valueMap.put("data", view.getDataList());
            valueMap.put("entity", view.getCls());
            exportParamList.add(valueMap);
        }
        // 执行方法
        Workbook workBook = ExcelExportUtil.exportExcel(exportParamList, XSSF);
        ExcelUtil.downloadExcel(response, workBook, "计算机二班选课情况");
    }



}
