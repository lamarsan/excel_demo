package com.lamarsan.excel_demo.model.actualcombat;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import java.util.List;

/**
 * className: ElectronicVersionModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:16
 */
@Data
public class ElectronicVersionModel {
    @Excel(name = "商务、开标")
    private List<ElectronicVersionNumModel> electronicVersionNumModels;
    @Excel(name = "技术")
    private List<TechnologyNumModel> technologyNumModels;
}
