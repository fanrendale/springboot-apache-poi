
package com.xjf.apachepoitest.util;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

/**
 * Excel处理类
 * @author xjf
 * @date 2019/2/26 9:40
 */
@Component
public class ExcelUtils {
    private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    public static Workbook loadExcel(MultipartFile retreatFile) {
        Workbook workbook = null;

        String fname = retreatFile.getOriginalFilename();
        String suffix = fname.substring(fname.lastIndexOf(".") + 1);
        if (!("xls".equals(suffix) || "xlsx".equals(suffix))){
            throw new RuntimeException("文件类型不正确");
        }

        try {
            if (fname.toLowerCase().endsWith(".xls")){
                return new HSSFWorkbook(retreatFile.getInputStream());
            }
            if (fname.toLowerCase().endsWith(".xlsx")){
                return new XSSFWorkbook(retreatFile.getInputStream());
            }
        } catch (IOException e) {
            throw new RuntimeException("文件加载异常");
        }

        return workbook;
    }

    public static Object getCellValue(Sheet sheet, int rowIndex, int cellIndex, Class cellValueType) {
        Row row = sheet.getRow(rowIndex);
        if (row == null){
            return null;
        }

        try {
            Cell cell = row.getCell(cellIndex);
            if (cell == null || cell.getCellTypeEnum() ==  CellType.BLANK){
                return null;
            }

            if (cellValueType == String.class) {
                return cell.getStringCellValue().replaceAll("\u00A0|\\s*|\r|\t|\n", "");
            } else if (cellValueType == Double.class) {
                return cell.getNumericCellValue();
            } else if (cellValueType == Date.class) {
                return cell.getDateCellValue();
            } else {
                return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    public static void exportTemplate(HttpServletRequest request, HttpServletResponse response, String fileName) {
        response.setContentType("application/vnd.ms-excel");
        // /remit-portal/src/main/webapp/WEB-INF/template/xxx.xlsx
        String nowPath = request.getSession().getServletContext().getRealPath("/") + "/" + "WEB-INF" + "/" + "template" + "/" + fileName;
        File file = new File(nowPath);
        // 清空response
        response.reset();
        // 设置response的Header
        OutputStream toClient = null;
        try {
            response.addHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes("gbk"), "iso-8859-1"));
            response.addHeader("Content-Length", "" + file.length());
            // 以流的形式下载文件
            InputStream fis = new BufferedInputStream(new FileInputStream(nowPath));
            byte[] buffer = new byte[fis.available()];
            fis.read(buffer);
            fis.close();

            toClient = new BufferedOutputStream(response.getOutputStream());
            toClient.write(buffer);
            toClient.flush();
            toClient.close();
        } catch (Exception e) {
            throw new RuntimeException("模板导出异常",e);
        } finally {
            if (toClient != null) {
                try {
                    toClient.close();
                } catch (IOException e) {
                    logger.info("", e);
                }

            }
        }

    }

    /**
     * 读取excel数据,返回List<map>,一个map就是一行数据,key为列名,value为列值(数据会四舍五入去整，注意...)
     * @param excelFile excel文件
     * @return List
     */
    public List<LinkedHashMap<String, String>> readExcel(MultipartFile excelFile,int sheetIndex) {
        logger.info("method=readExcel(),excelFile={}", excelFile.getOriginalFilename());
        try {
            // 效验文件格式
            String fileName = excelFile.getOriginalFilename();
            if (fileName == null || "".equals(fileName)){
                return null;
            }
            String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
            if (!("xls".equals(suffix) || "xlsx".equals(suffix))) {
                throw new RuntimeException("文件类型不正确,当前类型：【" + suffix + "】");
            }

            // 创建book对象
            Workbook wb = null;
            try {
                if (fileName.toLowerCase().endsWith(".xls")){
                    wb = create03WookBook(excelFile.getInputStream());
                }
                else if (fileName.toLowerCase().endsWith(".xlsx")){
                    wb = create07WookBook(excelFile.getInputStream());
                }
            } catch (Exception e) {
                logger.error("", e);
                throw new RuntimeException("文件为空或类型不正确,当前文件：【" + fileName + "】");
            }

            // 检查表格数据是否为空
            Sheet sheet = wb.getSheetAt(sheetIndex);
            if (sheet == null || sheet.getLastRowNum() < 1) {
                throw new RuntimeException("解析到文件【" + fileName + "】无数据");
            }

            // 删除无效行
            sheet = delInvalidRow(sheet);

            // 删除无效列
            sheet = delInvalidCol(sheet);

            Row row0 = sheet.getRow(0);
            // 检查是否有重复列名
            isRepeatCol(row0);

            List<LinkedHashMap<String, String>> list = new ArrayList<>();
            // 循环读取每行 每列 数据
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (null == row){
                    continue;
                }
                LinkedHashMap<String, String> map = new LinkedHashMap<>();
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    map.put(getStringValue(row0, j), getStringValue(row, j));
                }
                map.remove("");
                list.add(map);
            }
            logger.info("method=readExcel(),read data = " + list);
            return list;
        } catch (Exception e) {
            logger.error("", e);
            throw e;
        }
    }

    /**
     * @Description 获取每一格的数据 (统一返回为String类型)
     * @param row 行
     * @param cellIndex 表格下标
     * @return String 表格数据
     */
    public String getStringValue(Row row, int cellIndex) {
        Cell cell = row.getCell(cellIndex);
        String value = null;
        switch (cell.getCellTypeEnum()) {
            case  BLANK:
                break;
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) { // 日期类型
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
                    value = sdf.format(date);
                } else {
                    Double data = cell.getNumericCellValue();
                    value = data.toString();
                }
                break;
            case FORMULA:
                value = String.valueOf(cell.getNumericCellValue());
                if (value.equals("NaN")){
                    throw new RuntimeException("excel数字格式错误");
                }
            case ERROR:
                throw new RuntimeException("excel错误");
            default:
                break;
        }

        return value == null ? null : value.trim();

    }


    /** 删除无效(空白)行 */
    public Sheet delInvalidRow(Sheet sheet) {
        Row row0 = sheet.getRow(0);
        if (null == row0){
            throw new RuntimeException("第一行不能为空!");
        }
        boolean isEmpty = true;
        for (int i = 0; i < row0.getLastCellNum(); i++) {
            Cell cell = row0.getCell(i);
            if (null == cell || "".equals(getStringValue(row0, i))){
                continue;
            }
            isEmpty = false;
            break;
        }
        if (isEmpty){
            throw new RuntimeException("第一行不能为空!");
        }
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (null == row){
                continue;
            }

            boolean flag = true;
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (null == cell || "".equals(getStringValue(row, j))){
                    continue;
                }
                flag = false;
                break;
            }
            if (flag){
                sheet.removeRow(row);
            }
        }

        return sheet;
    }

    /** 删除无效(列名为空值)列 */
    public Sheet delInvalidCol(Sheet sheet) {
        Row row0 = sheet.getRow(0);
        for (int j = 0; j < row0.getLastCellNum(); j++) {
            String value = getStringValue(row0, j);
            if ("".equals(value)) {
                Cell cell = row0.getCell(j);
                if (cell != null){
                    row0.removeCell(cell);
                }
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (null == row){
                        continue;
                    }
                    Cell cell2 = row.getCell(j);
                    if (cell2 != null){
                        row.removeCell(cell2);
                    }
                }
            }
        }
        return sheet;
    }

    /**
     * 检查是否有重复列
     * @param row0 第一行(列名)
     */
    public void isRepeatCol(Row row0) {
        List<String> colList = new ArrayList<>();
        for (int i = 0; i < row0.getLastCellNum(); i++) {
            String val = getStringValue(row0, i);
            if (!"".equals(val)){
                colList.add(val);
            }
        }
        for (int i = 0; i < colList.size(); i++) {
            String val = colList.remove(i);
            if (colList.contains(val)){
                throw new RuntimeException("数据有误,重复列名:" + val);
            }
            i = -1;
        }
    }

    /**
     * @Description 创建03版的excel
     * @return Workbook
     * @param inputStream
     * @throws IOException
     */
    public Workbook create03WookBook(InputStream inputStream) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
        return workbook;
    }

    /**
     * @Description 创建07版的excel
     * @return Workbook
     * @param inputStream
     * @throws IOException
     */
    public Workbook create07WookBook(InputStream inputStream) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        return workbook;
    }
}
