package com.lamarsan.excel_demo.utils;

import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.text.DateFormat;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.regex.Pattern;

/**
 * className: ExcelUtil
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/11 16:14
 */
public class ExcelUtil {
    /**
     * 下载通用配置
     *
     * @param response  HttpServletResponse
     * @param workbook  Workbook
     * @param excelName excelName
     */
    public static void downloadExcel(HttpServletResponse response, Workbook workbook, String excelName) {
        try (OutputStream os = response.getOutputStream()) {

            response.reset();
            if (excelName == null) {
                excelName = UUID.randomUUID().toString();
            }
            response.setHeader("Content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "inline; filename=" + new String(excelName.getBytes(StandardCharsets.UTF_8), "ISO8859-1") + ".xlsx");

            workbook.write(os);
            os.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 复制上一行的格式样式
     * 
     * @param sheet
     * @param index
     * @param newRow
     * @return
     *     
     */
    public static Row setCellStype(Sheet sheet, int index, int newRow) {
        Row source = sheet.getRow(index);
        Row target = sheet.createRow(newRow);
        target.setHeight(source.getHeight());
        //target.setRowStyle(source.getRowStyle());
        for (int i = source.getFirstCellNum(); i < source.getLastCellNum(); i++) {
            Cell scell = source.getCell(i);
            Cell cell = target.createCell(i);
            cell.setCellStyle(scell.getCellStyle());
            cell.setCellType(scell.getCellType());
        }
        return target;
    }

    /**
     *   * 把sheet2的内容复制到sheet1中
     *   * 
     *   * @param sheet1
     *   * @param sheet2
     *  
     */
    public static void mergeSheet(Sheet sheet1, Sheet sheet2, int num1) {
        if (num1 == 0) {
            num1 = sheet1.getPhysicalNumberOfRows();
        }
        int num2 = sheet2.getPhysicalNumberOfRows();
        for (int i = 0; i < num2; i++) {
            Row row2 = sheet2.getRow(i);
            Row row1 = sheet1.createRow(num1 + 1 + i);
            row1.setHeight(row2.getHeight());
            row1.setRowStyle(row2.getRowStyle());
            int first = row2.getFirstCellNum();
            int last = row2.getLastCellNum();
            for (int j = first; j < last; j++) {
                Cell cell2 = row2.getCell(j);
                Cell cell1 = row1.createCell(j);
                cell1.setCellStyle(cell2.getCellStyle());
                cell1.setCellType(cell2.getCellType());
                cell1.setCellValue(cellString(cell2, ""));
            }

        }
    }

    /**
     *   * 取出cell里的值
     *   * 
     *   * @param cell
     *   * @param cellString
     *   * @return
     *  
     */
    public static String cellString(Cell cell, String cellString) {
        if (cell != null) {
            Object o = null;
            int cellType = cell.getCellType();
            switch (cellType) {
                case Cell.CELL_TYPE_BLANK:
                    o = "";
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    o = cell.getBooleanCellValue();
                    break;
                case Cell.CELL_TYPE_ERROR:
                    o = "Bad value!";
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    o = getValueOfNumericCell(cell);
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    try {
                        o = getValueOfNumericCell(cell);
                    } catch (IllegalStateException e) {
                        try {
                            o = cell.getRichStringCellValue().toString();
                        } catch (IllegalStateException e2) {
                            o = cell.getErrorCellValue();
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    break;
                default:
                    o = cell.getRichStringCellValue().getString();
            }
            cellString = String.valueOf(o);
        }
        return cellString;
    }

    private static Object getValueOfNumericCell(Cell cell) {
        Boolean isDate = DateUtil.isCellDateFormatted(cell);
        Double d = cell.getNumericCellValue();
        Object o = null;
        if (isDate) {
            o = DateFormat.getDateInstance().format(cell.getDateCellValue());
        } else {
            o = getRealStringValueOfDouble(d);
        }
        return o;
    }

    private static String getRealStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            // 小数部分
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint + BigInteger.ONE.intValue(), indexOfE));
            // 指数
            int pow = Integer.valueOf(doubleStr.substring(indexOfE + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            doubleStr = String.format("%." + scale + "f", d);
        } else {
            java.util.regex.Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }


    /**
     * 获取图片和位置 (xls)
     * @param sheet
     * @return
     * @throws IOException
     */
    public static Map<String, PictureData> getPictures1 (HSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = (HSSFClientAnchor) picture.getAnchor();
                PictureData pdata = picture.getPictureData();
                String key = cAnchor.getRow1() + "-" + cAnchor.getCol1(); // 行号-列号
                map.put(key, pdata);
            }
        }
        return map;
    }

    /**
     * 获取图片和位置 (xlsx)
     * @param sheet
     * @return
     * @throws IOException
     */
    public static Map<String, PictureData> getPictures2 (XSSFSheet sheet) throws IOException {
        Map<String, PictureData> map = new HashMap<String, PictureData>();
        List<POIXMLDocumentPart> list = sheet.getRelations();
        for (POIXMLDocumentPart part : list) {
            if (part instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) part;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker marker = anchor.getFrom();
                    String key = marker.getRow() + "-" + marker.getCol();
                    map.put(key, picture.getPictureData());
                }
            }
        }
        return map;
    }
    //图片写出
    public static void printImg(Map<String, PictureData> sheetList) throws IOException {

        //for (Map<String, PictureData> map : sheetList) {
        Object key[] = sheetList.keySet().toArray();
        for (int i = 0; i < sheetList.size(); i++) {
            // 获取图片流
            PictureData pic = sheetList.get(key[i]);
            // 获取图片索引
            String picName = key[i].toString();
            // 获取图片格式
            String ext = pic.suggestFileExtension();

            byte[] data = pic.getData();

            //图片保存路径
            FileOutputStream out = new FileOutputStream("D:\\记事本\\公司\\图片\\" + picName + "." + ext);
            out.write(data);
            out.close();
        }
        // }

    }
}
