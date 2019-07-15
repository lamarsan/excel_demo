package com.lamarsan.excel_demo.model.actualcombat;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * className: ActualCombatModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:08
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ActualCombatModel {
    @Excel(name = "序号",needMerge = true)
    private String id;
    @Excel(name = "投标人名称",needMerge = true)
    private String name;
    @Excel(name = "购买包号",needMerge = true)
    private String bagId;
    @Excel(name = "弃包",needMerge = true)
    private String bagName;
    //@ExcelCollection(name = "电子版（光盘）")
    //private List<ElectronicVersionModel> electronicVersionModels;
    //@ExcelCollection(name = "密封状况")
    //private List<SealingCondition> sealingConditions;
    @Excel(name = "递交人签字",needMerge = true)
    private String sign;
    @Excel(name = "联系电话",needMerge = true)
    private String phoneNum;
    //@ExcelCollection(name = "递交时间")
    //private List<TimeModel> timeModels;
    @Excel(name = "备注",needMerge = true)
    private String remark;
}
