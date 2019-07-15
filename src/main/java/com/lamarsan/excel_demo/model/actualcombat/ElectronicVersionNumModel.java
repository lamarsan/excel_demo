package com.lamarsan.excel_demo.model.actualcombat;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * className: ElectronicVersionNumModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:26
 */
@Data
public class ElectronicVersionNumModel {
    @Excel(name = "2份")
    private String name;
}
