package com.lamarsan.excel_demo.controller;

import com.lamarsan.excel_demo.model.PoiModel;
import com.lamarsan.excel_demo.service.PoiService;
import com.lamarsan.excel_demo.utils.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * className: PoiController
 * description: TODO
 *
 * @author hasee
 * @version 1.0
 * @date 2019/7/12 14:26
 */
@RestController
@RequestMapping("/excelPoi")
public class PoiController {

    @Autowired
    PoiService poiService;

    /**
     * 先导入Excel格式，然后填写内容，再导出
     * @param file 
     * @param response
     * @return
     */
    @PostMapping(value = "/excelImport")
    public Object upload(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        if (file == null) {
            System.out.println("文件不能为空");
        }
        List<PoiModel> list = new ArrayList<>();
        String fileName = file.getOriginalFilename(); //获取文件名
        try {
            HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(file.getInputStream()));
            // 有多少个sheet
            int sheets = workbook.getNumberOfSheets();
            for (int i = 0; i < sheets; i++) {
                //HSSFSheet sheet = workbook.getSheetAt(i);
                //// 获取多少行
                ////int rows = sheet.getPhysicalNumberOfRows();
                PoiModel poiModel = new PoiModel("1", "傅泽鹏", "1", "包名1", "1", "技术1", "是", "签字", "111", "2019", "无");
                PoiModel poiModel2 = new PoiModel("2", "徐镇涛", "2", "包名2", "2", "技术2", "是", "签字", "222", "2019", "无");
                PoiModel poiModel3 = new PoiModel("3", "李宁", "3", "包名3", "3", "技术3", "是", "签字", "333", "2019", "无");
                list.add(poiModel);
                list.add(poiModel2);
                list.add(poiModel3);
            }

        } catch (
                IOException e) {
            e.printStackTrace();
        }

        // 写入poi
        Workbook workBook = null;
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try {
            workBook = new XSSFWorkbook(file.getInputStream());
        } catch (
                Exception ex) {
            try {
                workBook = new HSSFWorkbook(file.getInputStream());
            } catch (Exception e) {

            }
        }
        try {
            // 写入数据
            if (workBook == null) {
                return "为空";
            }
            workBook = poiService.writerWorkbookContact(workBook, list);
            workBook.write(out);
            // 下载
            ExcelUtil.downloadExcel(response, workBook, fileName);
        } catch (
                Exception e) {
            e.printStackTrace();
            return "下载失败";
        }
        return list;
    }

    @PostMapping(value = "/importZip")
    private Object importZip(@RequestParam("file") MultipartFile zipFile) {
        //获得文件名
        String fileName = zipFile.getOriginalFilename();
        //检查文件
        if ("".equals(fileName)) {
            System.out.println("文件为空");
        }
        List<List<PoiModel>> poiModelLists = new ArrayList<>();
        try {
            //再本地创建一个文件，读取此文件 防止浏览器读取的文件被损坏
            File localFile = new File("D:\\记事本\\公司\\fyJyqdYhqdxxZip.zip");
            FileOutputStream ftpOutstream = new FileOutputStream(localFile);
            byte[] appByte = zipFile.getBytes();
            ftpOutstream.write(appByte);
            ftpOutstream.flush();
            ftpOutstream.close();//创建完毕后删除

            File file = new File("D:\\记事本\\公司\\fyJyqdYhqdxxZip.zip");
            //不解压直接读取,加上UTF-8解决乱码问题,file转ZipInputStream
            ZipInputStream in = new ZipInputStream(new FileInputStream(file), Charset.forName("GBK"));
            //不解压直接读取,加上UTF-8解决乱码问题,ZipInputStream转BufferedReader
            BufferedReader br = new BufferedReader(new InputStreamReader(in, "gbk"));
            //把InputStream转成ByteArrayOutputStream 多次使用
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            ZipEntry ze;
            while ((ze = in.getNextEntry()) != null) {
                if (ze.isDirectory()) {
                    //如果是目录，不处理
                    continue;
                }
                try {
                    String zipFileName = ze.getName();
                    //不是我们指定的文件不导入，XXXXX.市场化清单.xls
                    //if (zipFileName != null && zipFileName.indexOf(".") != -1
                    //        && zipFileName.equals(zipFileName.substring(0, zipFileName.indexOf(".xls")) + "市场化清单.xls")) {
                    //    continue;
                    //}
                    byte[] buffer = new byte[1024];
                    int len;
                    while ((len = in.read(buffer)) > -1) {
                        baos.write(buffer, 0, len);
                    }
                    baos.flush();

                    InputStream stream = new ByteArrayInputStream(baos.toByteArray());
                    //获取Excel对象
                    HSSFWorkbook wb = new HSSFWorkbook(stream);
                    int sheets = wb.getNumberOfSheets();
                    for (int i = 0; i < sheets; i++) {
                        HSSFSheet sheet = wb.getSheetAt(i);
                        // 获取多少行
                        List<PoiModel> poiModels = new ArrayList<>();
                        int rows = sheet.getPhysicalNumberOfRows();
                        for (int j = 0; j < rows; j++) {
                            //获取Row对象
                            HSSFRow row = sheet.getRow(j);
                            //获取Cell对象的值并输出
                            PoiModel poiModel = new PoiModel(row.getCell(0).toString(), row.getCell(1).toString(), row.getCell(2).toString(), row.getCell(3).toString(), row.getCell(4).toString(), row.getCell(5).toString(), row.getCell(6).toString(), row.getCell(7).toString(), row.getCell(8).toString(), row.getCell(9).toString(), row.getCell(10).toString());
                            System.out.println(row.getCell(0) + " " + row.getCell(1));
                            poiModels.add(poiModel);
                        }
                        poiModelLists.add(poiModels);
                    }
                    baos.reset();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            br.close();
            in.close();
            baos.close();
            //处理完毕删除
            localFile.delete();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return poiModelLists;
    }

    @PostMapping(value = "/excelImgFile")
    public Object getImg(@RequestParam("file") MultipartFile file, HttpServletResponse response) {
        if (file == null) {
            System.out.println("文件不能为空");
        }
        List<PoiModel> list = new ArrayList<>();
        String fileName = file.getOriginalFilename();
        if (!fileName.endsWith(".xls") && !fileName.endsWith(".xlsx")) {
            System.out.println("文件不是excel类型");
        }
        Workbook workbook = null;
        Sheet sheet;
        try {
            workbook = new HSSFWorkbook(new POIFSFileSystem(file.getInputStream()));
        } catch (IOException e) {
            e.printStackTrace();
            try {
                workbook = new XSSFWorkbook(file.getInputStream());
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        }
        Map<String, PictureData> mapList = null;

        sheet = workbook.getSheetAt(0);
        // 判断用07还是03的方法获取图片
        try {
            if (fileName.endsWith(".xls")) {
                mapList = ExcelUtil.getPictures1((HSSFSheet) sheet);
            } else if (fileName.endsWith(".xlsx")) {
                mapList = ExcelUtil.getPictures2((XSSFSheet) sheet);
            }
        } catch (IOException e) {
            System.out.println("不能获取图片");
        }

        try {
            ExcelUtil.printImg(mapList);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        List<PoiModel> poiModelList = poiService.getImg(workbook, sheet);
        return poiModelList;
    }
}
