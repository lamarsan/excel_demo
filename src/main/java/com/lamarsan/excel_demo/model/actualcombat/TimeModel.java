package com.lamarsan.excel_demo.model.actualcombat;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * className: TimeModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 11:32
 */
@Data
public class TimeModel {
    @Excel(name = "（2018年）")
    private String time;
}
