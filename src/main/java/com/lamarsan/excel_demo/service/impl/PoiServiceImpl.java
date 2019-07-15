package com.lamarsan.excel_demo.service.impl;

import com.lamarsan.excel_demo.model.PoiModel;
import com.lamarsan.excel_demo.service.PoiService;
import com.lamarsan.excel_demo.utils.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * className: PoiServiceImpl
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 14:53
 */
@Service
public class PoiServiceImpl implements PoiService {
    @Override
    public Workbook writerWorkbookContact(Workbook workBook, List<PoiModel> list) {
        Sheet sheet = workBook.getSheetAt(0);
        int size = list.size();
        //图片操作
        BufferedImage bufferImg = null;//图片一
        BufferedImage bufferImg2 = null;//图片二
        BufferedImage bufferImg3 = null;//图片二
        //图片
        try {
            // 先把读进来的图片放到一个ByteArrayOutputStream中，以便产生ByteArray
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            ByteArrayOutputStream byteArrayOut2 = new ByteArrayOutputStream();
            ByteArrayOutputStream byteArrayOut3 = new ByteArrayOutputStream();
            //将图片读到BufferedImage
            bufferImg = ImageIO.read(new File("D:\\图片\\140.png"));
            bufferImg2 = ImageIO.read(new File("D:\\图片\\137.png"));
            bufferImg3 = ImageIO.read(new File("D:\\图片\\139.png"));
            // 将图片写入流中
            ImageIO.write(bufferImg, "png", byteArrayOut);
            ImageIO.write(bufferImg2, "png", byteArrayOut2);
            ImageIO.write(bufferImg3, "png", byteArrayOut3);
            // 利用HSSFPatriarch将图片写入EXCEL
            Drawing patriarch = sheet.createDrawingPatriarch();
            /**
             * 该构造函数有8个参数
             * 前四个参数是控制图片在单元格的位置，分别是图片距离单元格left，top，right，bottom的像素距离
             * 后四个参数，前两个表示图片左上角所在的cellNum和 rowNum，后两个参数对应的表示图片右下角所在的cellNum和 rowNum，
             * excel中的cellNum和rowNum的index都是从0开始的
             *
             */
            //图片一导出到单元格B2中
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                    (short) 7, 4, (short) 8, 5);
            HSSFClientAnchor anchor2 = new HSSFClientAnchor(0, 0, 0, 0,
                    (short) 7, 5, (short) 8, 6);
            HSSFClientAnchor anchor3 = new HSSFClientAnchor(0, 0, 0, 0,
                    (short) 7, 6, (short) 8, 7);
            // 插入图片
            patriarch.createPicture(anchor, workBook.addPicture(byteArrayOut
                    .toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            patriarch.createPicture(anchor2, workBook.addPicture(byteArrayOut2
                    .toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            patriarch.createPicture(anchor3, workBook.addPicture(byteArrayOut3
                    .toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
        } catch (IOException io) {
            io.printStackTrace();
            System.out.println("io erorr : " + io.getMessage());
        }
        for (int i = 0; i < size; i++) {
            PoiModel c = list.get(i);
            Row row = null;
            if (i == 0) {
                row = sheet.getRow(4);
            } else {
                row = ExcelUtil.setCellStype(sheet, 4, 4 + i);
            }
            row.getCell(0).setCellValue(c.getId());
            row.getCell(1).setCellValue(c.getName());
            row.getCell(2).setCellValue(c.getBagId());
            row.getCell(3).setCellValue(c.getBagName());
            row.getCell(4).setCellValue(c.getElectronicVersionName());
            row.getCell(5).setCellValue(c.getTechnology());
            row.getCell(6).setCellValue(c.getSealingState());
            row.getCell(7).setCellValue(c.getSign());
            row.getCell(8).setCellValue(c.getPhoneNum());
            row.getCell(9).setCellValue(c.getTime());
            row.getCell(10).setCellValue(c.getRemark());
        }

        // 合并单元格
        //sheet.addMergedRegion(new CellRangeAddress(size + 5, size + 16, 14, 15));
        //sheet.addMergedRegion(new CellRangeAddress(size + 23, size + 23, 0, 5));
        //sheet.addMergedRegion(new CellRangeAddress(size + 27, size + 27, 0, 5));
        //sheet.addMergedRegion(new CellRangeAddress(size + 32, size + 32, 0, 5));
        return workBook;
    }

    @Override
    public List<PoiModel> getImg(Workbook workBook, Sheet sheet) {
        //获得表头
        Row rowHead = sheet.getRow(0);

        //判断表头是否正确
        System.out.println(rowHead.getPhysicalNumberOfCells());
        if (rowHead.getPhysicalNumberOfCells() != 11) {
            System.out.println("表头的数量不对!");
        }

        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();
        List<PoiModel> poiModelList = new ArrayList<>();
        //获得所有数据
        for (int i = 1; i <= totalRowNum; i++) {
            PoiModel poiModel = new PoiModel();
            //获得第i行对象
            Row row = sheet.getRow(i);
            //获得获得第i行第0列的 String类型对象
            Cell cell = row.getCell((short) 0);
            poiModel.setId(cell.getStringCellValue());
            cell = row.getCell((short) 1);
            poiModel.setName(cell.getStringCellValue());
            cell = row.getCell((short) 2);
            poiModel.setBagId(cell.getStringCellValue());
            cell = row.getCell((short) 3);
            poiModel.setElectronicVersionName(cell.getStringCellValue());
            cell = row.getCell((short) 4);
            poiModel.setTechnology(cell.getStringCellValue());
            cell = row.getCell((short) 5);
            poiModel.setSealingState(cell.getStringCellValue());
            cell = row.getCell((short) 7);
            poiModel.setPhoneNum(cell.getStringCellValue());
            cell = row.getCell((short) 8);
            poiModel.setTime(cell.getStringCellValue());
            cell = row.getCell((short) 9);
            poiModel.setRemark(cell.getStringCellValue());
            poiModelList.add(poiModel);
        }
        return poiModelList;
    }
}
