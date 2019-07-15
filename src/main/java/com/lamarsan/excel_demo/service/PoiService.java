package com.lamarsan.excel_demo.service;

import com.lamarsan.excel_demo.model.PoiModel;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * className: PoiService
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 14:51
 */
public interface PoiService {
    /**
     * 写内容
     *
     * @param workBook
     * @param list
     * @return
     */
    Workbook writerWorkbookContact(Workbook workBook, List<PoiModel> list);

    /**
     * 获取信息
     * @param workBook
     * @param sheet
     */
    List<PoiModel> getImg(Workbook workBook, Sheet sheet);
}
