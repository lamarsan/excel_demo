package com.lamarsan.excel_demo.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * className: PoiModel
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 15:56
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class PoiModel {
    private String id;
    private String name;
    private String bagId;
    private String bagName;
    private String electronicVersionName;
    private String technology;
    private String sealingState;
    private String sign;
    private String phoneNum;
    private String time;
    private String remark;
}
