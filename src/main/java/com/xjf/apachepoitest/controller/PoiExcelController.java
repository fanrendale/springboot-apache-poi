package com.xjf.apachepoitest.controller;

import com.xjf.apachepoitest.util.ExcelUtils;
import io.swagger.annotations.ApiOperation;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * Apache Poi处理Excel
 * @author xjf
 * @date 2019/2/26 9:40
 */
@RestController
public class PoiExcelController {
    @Autowired
    private ExcelUtils excelUtils;

    /**
     * 在桌面上生成一个Excel文件
     */
    @ApiOperation(value = "生成一个Excel文件",notes = "在桌面生成一个Excel文件")
    @GetMapping("/createExcel")
    public  String  createExcel() throws IOException {
        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        String filePath = desktop + "/template.xls";

        File file = new File(filePath);
        OutputStream outputStream = new FileOutputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Sheet1");
        HSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("id");
        row.createCell(1).setCellValue("订单号");
        row.createCell(2).setCellValue("下单时间");
        row.createCell(3).setCellValue("个数");
        row.createCell(4).setCellValue("单价");
        row.createCell(5).setCellValue("订单金额");
        // 设置行的高度
        row.setHeightInPoints(30);

        HSSFRow row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("1");
        row1.createCell(1).setCellValue("NO00001");

        // 日期格式化
        HSSFCellStyle cellStyle2 = workbook.createCellStyle();
        HSSFCreationHelper creationHelper = workbook.getCreationHelper();
        cellStyle2.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
        // 设置列的宽度
        sheet.setColumnWidth(2, 20 * 256);

        HSSFCell cell2 = row1.createCell(2);
        cell2.setCellStyle(cellStyle2);
        cell2.setCellValue(new Date());

        row1.createCell(3).setCellValue(2);


        // 保留两位小数
        HSSFCellStyle cellStyle3 = workbook.createCellStyle();
        cellStyle3.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
        HSSFCell cell4 = row1.createCell(4);
        cell4.setCellStyle(cellStyle3);
        cell4.setCellValue(29.5);


        // 货币格式化
        HSSFCellStyle cellStyle4 = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setFontName("华文行楷");
        font.setFontHeightInPoints((short)15);
        font.setColor(HSSFColor.RED.index);
        cellStyle4.setFont(font);

        HSSFCell cell5 = row1.createCell(5);
        // 设置计算公式
        cell5.setCellFormula("D2*E2");

        // 获取计算公式的值
        HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(workbook);
        cell5 = e.evaluateInCell(cell5);
        System.out.println(cell5.getNumericCellValue());


        workbook.setActiveSheet(0);
        workbook.write(outputStream);
        outputStream.close();

        return "create success!";
    }

    /**
     * 读取Excel，解析数据
     * @throws IOException
     */
    @ApiOperation(value = "读取Excel文件",notes = "读取桌面的Excel文件")
    @GetMapping("/readExcel")
    public String readExcel() throws IOException{
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        String filePath = desktop + "/template.xls";

        FileInputStream fileInputStream = new FileInputStream(filePath);
        BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
        POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
        HSSFWorkbook workbook = new HSSFWorkbook(fileSystem);
        HSSFSheet sheet = workbook.getSheet("Sheet1");

        int lastRowIndex = sheet.getLastRowNum();
        System.out.println(lastRowIndex);
        for (int i = 0; i <= lastRowIndex; i++) {
            HSSFRow row = sheet.getRow(i);
            if (row == null) { break; }

            short lastCellNum = row.getLastCellNum();
            DataFormatter formatter = new DataFormatter();

            for (int j = 0; j < lastCellNum; j++) {

                String cellValue = getHssfCellValue(row.getCell(j));
                System.out.println(cellValue);
            }
        }


        bufferedInputStream.close();

        return "read success";
    }

    /**
     * 使用工具类读取Excel
     * @param file
     * @return
     */
    @ApiOperation(value = "使用工具类读取Excel",notes = "使用工具类读取Excel")
    @PostMapping("/readExcelByUtil")
    public String readExcelByUtil(@RequestParam("file")MultipartFile file){
        List<LinkedHashMap<String, String>> excelInfo = excelUtils.readExcel(file,0);

        excelInfo.forEach(info->{
            System.out.println("id:"+ info.get("id"));
            System.out.println("订单号:"+ info.get("订单号"));
            System.out.println("下单时间:"+ info.get("下单时间"));
            System.out.println("个数:"+ info.get("个数"));
            System.out.println("单价:"+ info.get("单价"));
            System.out.println("订单金额:"+ info.get("订单金额"));

        });

        return "read success";
    }

    /**
     * 获取单元格数据
     * @param hssfCell
     * @return
     */
    private String getHssfCellValue(HSSFCell hssfCell) {
        String cellvalue="";
        DataFormatter formatter = new DataFormatter();
        if (null != hssfCell) {
            switch (hssfCell.getCellType()) {
                case HSSFCell.CELL_TYPE_NUMERIC: // 数字
                    if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(hssfCell)) {
                        cellvalue = formatter.formatCellValue(hssfCell);
                    } else {
                        double value = hssfCell.getNumericCellValue();
                        int intValue = (int) value;
                        cellvalue = value - intValue == 0 ? String.valueOf(intValue) : String.valueOf(value);
                    }
                    break;
                case HSSFCell.CELL_TYPE_STRING: // 字符串
                    cellvalue=hssfCell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
                    cellvalue=String.valueOf(hssfCell.getBooleanCellValue());
                    break;
                case HSSFCell.CELL_TYPE_FORMULA: // 公式
                    cellvalue=String.valueOf(hssfCell.getCellFormula());
                    break;
                case HSSFCell.CELL_TYPE_BLANK: // 空值
                    cellvalue="";
                    break;
                case HSSFCell.CELL_TYPE_ERROR: // 故障
                    cellvalue="";
                    break;
                default:
                    cellvalue="UNKNOWN TYPE";
                    break;
            }
        } else {
            System.out.print("-");
        }
        return cellvalue.trim();
    }
}
