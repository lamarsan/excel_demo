package com.lamarsan.excel_demo.controller;

import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import com.lamarsan.excel_demo.model.actualcombat.ActualCombatModel;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;

/**
 * className: ExcelActualCombatController
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:05
 */
@RestController
@RequestMapping("/excel")
public class ExcelActualCombatController {

    /**
     * excel导入导出
     *
     */
    @PostMapping(value = "/excelImport")
    public Object importExcel(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        //接收导入数组
        List<ActualCombatModel> actualCombatModels = null;
        try {
            ImportParams params = new ImportParams();
            //标题行号
            params.setTitleRows(1);
            //开始的行数
            params.setStartRows(2);
            actualCombatModels = ExcelImportUtil.importExcel(file.getInputStream(), ActualCombatModel.class, params);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return actualCombatModels;
    }
}
