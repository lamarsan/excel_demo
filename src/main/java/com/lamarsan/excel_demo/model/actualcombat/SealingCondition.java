package com.lamarsan.excel_demo.model.actualcombat;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * className: SealingCondition
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:21
 */
@Data
public class SealingCondition {
    @Excel(name = "是否完好")
    private String state;
}
