package com.lamarsan.excel_demo.common;

import com.google.common.collect.Lists;
import lombok.AllArgsConstructor;

import java.util.List;

/**
 * className: ExportMoreView
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 19:01
 */
@AllArgsConstructor
public class ExportMoreView {
    private List<ExportView> moreViewList= Lists.newArrayList();

    public List<ExportView> getMoreViewList() {
        return moreViewList;
    }

    public void setMoreViewList(List<ExportView> moreViewList) {
        this.moreViewList = moreViewList;
    }
}
