package com.framework.excel;


import com.framework.config.SystemConfigure;
import com.framework.enums.CantExcelMultilineEnum;
import com.framework.excel.model.*;
import com.framework.log.SystemLogUtil;
import com.framework.util.CollectionUtil;
import com.framework.util.DateUtil;
import com.framework.util.StringUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.IOUtils;
import org.apache.http.client.utils.DateUtils;
import org.apache.logging.log4j.util.Strings;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.Color;
import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.net.URLEncoder;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BiFunction;
import java.util.function.Function;

@Component
public class ExcelUtils {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static final String WORD_DOC = "doc";
    private static final String WORD_DOCX = "docx";
    public static BiFunction<Object, Object, String> jsbl = (a, b) -> {
        try {

            BigDecimal big1 = new BigDecimal(0);

            BigDecimal big2 = new BigDecimal(0);

            if (a != null || b != null) {

                if (a instanceof Long) {
                    big1 = new BigDecimal((long) a);
                } else if (a instanceof Double) {
                    big1 = new BigDecimal((double) a);
                } else if (a instanceof Integer) {
                    big1 = new BigDecimal((int) a);
                } else if (a instanceof BigDecimal) {
                    big1 = (BigDecimal) a;
                } else if (a instanceof String) {
                    big1 = new BigDecimal((String) a);
                } else {
                    return "0";
                }


                if (b instanceof Long) {
                    big2 = new BigDecimal((long) b);
                } else if (b instanceof Double) {
                    big2 = new BigDecimal((double) b);
                } else if (b instanceof Integer) {
                    big2 = new BigDecimal((int) b);
                } else if (b instanceof BigDecimal) {
                    big2 = (BigDecimal) b;
                } else if (b instanceof String) {
                    big2 = new BigDecimal((String) b);
                } else {
                    return "0";
                }


                if (big2.doubleValue() == 0) {
                    return "0";
                } else {
                    return big1.divide(big2, 2, RoundingMode.HALF_UP).toString();
                }
            } else {
                return "0";
            }


        } catch (Exception e) {

            System.out.println("分子 分母 错误！！" + e.toString());
            return "0";
        }
    };
    /**
     * 生成空白填充格
     * 传入多少个空白填充单元格
     */
    public static Function<Integer, List<String>> fillSpare = ln -> {
        List<String> stringList = new ArrayList<>();
        for (int i = 0; i < ln; i++) {
            stringList.add("");
        }

        return stringList;
    };
    /**
     * 生成空白填充格
     * 传入多少个空白填充单元格
     */
    public static Function<Integer, List<ExclMultilieRows>> fillSpareMultilie = ln -> {
        List<ExclMultilieRows> exclMultilieRowsList = new ArrayList<>();
        for (int i = 0; i < ln; i++) {
            exclMultilieRowsList.add(new ExclMultilieRows(""));
        }

        return exclMultilieRowsList;
    };
    /**
     * 每列宽度
     * 传入Integer数值，包含每列宽度
     */
    public static Function<Integer[], List<Integer>> colWidths = integers -> {
        List<Integer> colwidth = new ArrayList<>();
        for (int i = 0; i < integers.length; i++) {
            colwidth.add(integers[i]);
        }

        return colwidth;
    };
    protected final org.slf4j.Logger logger = LoggerFactory.getLogger(getClass());

    public static void exportExcel(HttpServletResponse response, String fileName, List<ExcelData> dataList) throws Exception {
        // 告诉浏览器用什么软件可以打开此文件
        response.setHeader("content-Type", "application/vnd.ms-excel");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "utf-8"));

        exportExcel(dataList, response.getOutputStream());
    }

    public static void exportExcel(List<ExcelData> dataList, OutputStream out) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook();
        try {

            for (ExcelData data : dataList) {
                String sheetName = data.getName();
                if (null == sheetName) {
                    sheetName = "Sheet1";
                }
                XSSFSheet sheet = wb.createSheet(sheetName);
                writeExcel(wb, sheet, data);
            }
            wb.write(out);
        } finally {
            wb.close();
        }
    }

    private static void writeExcel(XSSFWorkbook wb, Sheet sheet, ExcelData data) {

        int rowIndex = 0;
        if (CollectionUtil.isNotBlank(data.getHead())) {
            rowIndex = writeHeradToExcel(wb, sheet, data.getHead(), data.getHeadStyle(), data.getHeadMergedRegions());
        }
        if (CollectionUtil.isNotBlank(data.getTitles())) {
            if (data.getIs_MultiTitles() == CantExcelMultilineEnum.NO.getCode())

            {
                rowIndex = writeTitlesToExcel(wb, sheet, data.getTitles(), data.getTitleStyle(), rowIndex, data.getTitlesMergedRegions());
            } else {
                rowIndex = writeMultilieTitlesToExcel(wb, sheet, data.getTitles(), data.getTitleStyle(), rowIndex);
            }
        }
        //

        if (CollectionUtil.isNotBlank(data.getRows())) {
            if (data.getIs_Multiline() == CantExcelMultilineEnum.NO.getCode()) {
                writeRowsToExcel(wb, sheet, data.getRows(), data.getRowStyle(), rowIndex, data.getRowsMergedRegions());
            } else {
                writeMultilieRowsToExcel(wb, sheet, data.getRows(), data.getRowStyle(), rowIndex, data.getRowsMergedRegions());
            }
        }

        if (data.isIs_autoWith()) {
            if (CollectionUtil.isNotBlank(data.getTitles())) {
                autoSizeColumns(sheet, data.getTitles().size() + 1);
            } else {
                autoSizeColumns(sheet, data.getColSize() + 1);
            }
        }

        if (CollectionUtil.isNotBlank(data.getColsWidth())) {
            setSizeColumns(sheet, data.getColsWidth());
        }

        if (data.getColsHight() > 0) {
            sheet.setDefaultRowHeight(data.getColsHight());
        }

    }

    private static int writeHeradToExcel(XSSFWorkbook wb, Sheet sheet, List<List<String>> heads, List<ExcelFontStyle> headsStyle, List<ExcelMergedRegion> headMerged) {
        int rowIndex = 0;
        int colIndex = 0;


//        titleStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//        titleStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);


        for (int i = 0; i < heads.size(); i++) {
            colIndex = 0;
            Font titleFont = wb.createFont();
            XSSFCellStyle titleStyle = wb.createCellStyle();

            if (CollectionUtil.isNotBlank(headsStyle) && i < headsStyle.size()) {
                if (headsStyle.get(i).getFontName() != null) {
                    titleFont.setFontName(headsStyle.get(i).getFontName());
                } else {
                    titleFont.setFontName("宋体");
                }

                titleFont.setBold(headsStyle.get(i).isBold());

                if (headsStyle.get(i).getFontHeightInPoints() != 0) {
                    titleFont.setFontHeightInPoints(headsStyle.get(i).getFontHeightInPoints());
                } else {
                    titleFont.setFontHeightInPoints((short) 9);
                }

                if (headsStyle.get(i).getColor() != 0) {
                    titleFont.setColor(headsStyle.get(i).getColor());
                } else {
                    titleFont.setColor(IndexedColors.BLACK.index);
                }

                if (headsStyle.get(i).getHorizontalAlignment() != null) {
                    titleStyle.setAlignment(headsStyle.get(i).getHorizontalAlignment());
                } else {
                    titleStyle.setAlignment(HorizontalAlignment.CENTER);
                }

                if (headsStyle.get(i).getVerticalAlignment() != null) {
                    titleStyle.setVerticalAlignment(headsStyle.get(i).getVerticalAlignment());
                } else {
                    titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                }

                if (headsStyle.get(i).getFillForegroundColor() != null) {
                    titleStyle.setFillForegroundColor(new XSSFColor(headsStyle.get(i).getFillForegroundColor()));
                } else {
                    titleStyle.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255)));
                }

                if (headsStyle.get(i).getFillPatternType() != null) {
                    titleStyle.setFillPattern(headsStyle.get(i).getFillPatternType());
                } else {
                    titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }


                titleStyle.setWrapText(headsStyle.get(i).isWrapText());
                titleStyle.setFont(titleFont);
                if (headsStyle.get(i).isBorder()) {
                    Color color = Optional.ofNullable(headsStyle.get(i).getBorderColor()).orElse(new Color(0, 0, 0));
                    BorderStyle borderStyle = Optional.ofNullable(headsStyle.get(i).getBorderStyle()).orElse(BorderStyle.THIN);

                    setBorder(titleStyle, borderStyle, new XSSFColor(color));
                } else {

                    setBorder(titleStyle, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));
                }


            } else {
                // titleFont.setFontName("simsun");
                titleFont.setFontName("宋体");
                titleFont.setBold(true);
                titleFont.setFontHeightInPoints((short) 9);
                titleFont.setColor(IndexedColors.BLACK.index);

                titleStyle.setAlignment(HorizontalAlignment.CENTER);
                titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                titleStyle.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255)));
                //titleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                titleStyle.setFont(titleFont);
                setBorder(titleStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));
            }


            Row titleRow = sheet.createRow(rowIndex);
            // titleRow.setHeightInPoints(25);
            List<String> rowData = heads.get(i);


            for (String field : rowData) {


                Cell cell = titleRow.createCell(colIndex);
                cell.setCellValue(field);
                cell.setCellStyle(titleStyle);

                colIndex++;
            }


            rowIndex++;
            // mergedRegion(sheet,mergeCells[i],1,sheet.getLastRowNum(),workbook,mergeBasis);
        }

        for (int j = 0; j < headMerged.size(); j++) {
            sheet.addMergedRegion(new CellRangeAddress(headMerged.get(j).getFirstRow(), headMerged.get(j).getLastRow(), headMerged.get(j).getFirstCol(), headMerged.get(j).getLastCol()));
        }

        return rowIndex;

    }

    private static int writeTitlesToExcel(XSSFWorkbook wb, Sheet sheet, List<List<String>> titles, List<ExcelFontStyle> titleSytle, int rowIndex, List<ExcelMergedRegion> titleMerged) {
        if (CollectionUtil.isBlank(titles)) {
            return rowIndex;
        }
        // int rowIndex = 0;
        int colIndex = 0;
        int startIndex = rowIndex;
        for (int i = 0; i < titles.size(); i++) {

            Font titleFont = wb.createFont();
            XSSFCellStyle titleStyle1 = wb.createCellStyle();
            if (CollectionUtil.isNotBlank(titleSytle) && i < titleSytle.size()) {
                if (titleSytle.get(i).getFontName() != null) {
                    titleFont.setFontName(titleSytle.get(i).getFontName());
                } else {
                    titleFont.setFontName("宋体");
                }

                titleFont.setBold(titleSytle.get(i).isBold());

                if (titleSytle.get(i).getFontHeightInPoints() != 0) {
                    titleFont.setFontHeightInPoints(titleSytle.get(i).getFontHeightInPoints());
                } else {
                    titleFont.setFontHeightInPoints((short) 9);
                }

                if (titleSytle.get(i).getColor() != 0) {
                    titleFont.setColor(titleSytle.get(i).getColor());
                } else {
                    titleFont.setColor(IndexedColors.BLACK.index);
                }

                if (titleSytle.get(i).getHorizontalAlignment() != null) {
                    titleStyle1.setAlignment(titleSytle.get(i).getHorizontalAlignment());
                } else {
                    titleStyle1.setAlignment(HorizontalAlignment.CENTER);
                }

                if (titleSytle.get(i).getVerticalAlignment() != null) {
                    titleStyle1.setVerticalAlignment(titleSytle.get(i).getVerticalAlignment());
                } else {
                    titleStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
                }

                if (titleSytle.get(i).getFillForegroundColor() != null) {
                    titleStyle1.setFillForegroundColor(new XSSFColor(titleSytle.get(i).getFillForegroundColor()));
                } else {
                    titleStyle1.setFillForegroundColor(new XSSFColor(new Color(230, 230, 230)));
                }

                if (titleSytle.get(i).getFillPatternType() != null) {
                    titleStyle1.setFillPattern(titleSytle.get(i).getFillPatternType());
                } else {
                    titleStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }


                titleStyle1.setWrapText(titleSytle.get(i).isWrapText());
                titleStyle1.setFont(titleFont);

                if (titleSytle.get(i).isBorder()) {
                    Color color = Optional.ofNullable(titleSytle.get(i).getBorderColor()).orElse(new Color(0, 0, 0));
                    BorderStyle borderStyle = Optional.ofNullable(titleSytle.get(i).getBorderStyle()).orElse(BorderStyle.THIN);

                    setBorder(titleStyle1, borderStyle, new XSSFColor(color));
                } else {

                    setBorder(titleStyle1, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));
                }

            } else {


                titleFont.setFontName("simsun");
                titleFont.setBold(true);
                titleFont.setFontHeightInPoints((short) 9);
                titleFont.setColor(IndexedColors.BLACK.index);


//        titleStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//        titleStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                titleStyle1.setAlignment(HorizontalAlignment.CENTER);
                titleStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
                titleStyle1.setFillForegroundColor(new XSSFColor(new Color(230, 230, 230)));
                //titleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                titleStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                titleStyle1.setFont(titleFont);
                setBorder(titleStyle1, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));
            }
            Row titleRow = sheet.createRow(rowIndex);
            // titleRow.setHeightInPoints(25);
            colIndex = 0;

            //  int titleLine = 0;

            colIndex = 0;
            List<String> fieldList = titles.get(i);
            for (String field : fieldList) {
                Cell cell = titleRow.createCell(colIndex);
                cell.setCellValue(field);
                titleStyle1.setWrapText(true);
                cell.setCellStyle(titleStyle1);
                colIndex++;
            }


            if (CollectionUtil.isNotBlank(titleMerged) && i <= titleMerged.size()) {
                sheet.addMergedRegion(new CellRangeAddress(titleMerged.get(i).getFirstRow() + startIndex, titleMerged.get(i).getLastRow() + startIndex, titleMerged.get(i).getFirstCol(), titleMerged.get(i).getLastCol()));
            }
            // titleLine++;
            rowIndex++;
        }


        return rowIndex;
    }

    private static int writeMultilieTitlesToExcel(XSSFWorkbook wb, Sheet sheet, List<List<ExclMultilieRows>> titles, List<ExcelFontStyle> titleSytle, int rowIndex) {
        if (CollectionUtil.isBlank(titles)) {
            return rowIndex;
        }

        // int rowIndex = 0;
        int colIndex = 0;

        for (int i = 0; i < titles.size(); i++) {

            Font titleFont = wb.createFont();
            XSSFCellStyle titleStyle1 = wb.createCellStyle();
            if (CollectionUtil.isNotBlank(titleSytle) && i < titleSytle.size()) {
                if (titleSytle.get(i).getFontName() != null) {
                    titleFont.setFontName(titleSytle.get(i).getFontName());
                } else {
                    titleFont.setFontName("宋体");
                }

                titleFont.setBold(titleSytle.get(i).isBold());

                if (titleSytle.get(i).getFontHeightInPoints() != 0) {
                    titleFont.setFontHeightInPoints(titleSytle.get(i).getFontHeightInPoints());
                } else {
                    titleFont.setFontHeightInPoints((short) 9);
                }

                if (titleSytle.get(i).getColor() != 0) {
                    titleFont.setColor(titleSytle.get(i).getColor());
                } else {
                    titleFont.setColor(IndexedColors.BLACK.index);
                }

                if (titleSytle.get(i).getHorizontalAlignment() != null) {
                    titleStyle1.setAlignment(titleSytle.get(i).getHorizontalAlignment());
                } else {
                    titleStyle1.setAlignment(HorizontalAlignment.CENTER);
                }

                if (titleSytle.get(i).getVerticalAlignment() != null) {
                    titleStyle1.setVerticalAlignment(titleSytle.get(i).getVerticalAlignment());
                } else {
                    titleStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
                }

                if (titleSytle.get(i).getFillForegroundColor() != null) {
                    titleStyle1.setFillForegroundColor(new XSSFColor(titleSytle.get(i).getFillForegroundColor()));
                } else {
                    titleStyle1.setFillForegroundColor(new XSSFColor(new Color(230, 230, 230)));
                }

                if (titleSytle.get(i).getFillPatternType() != null) {
                    titleStyle1.setFillPattern(titleSytle.get(i).getFillPatternType());
                } else {
                    titleStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }


                titleStyle1.setWrapText(titleSytle.get(i).isWrapText());
                titleStyle1.setFont(titleFont);

                if (titleSytle.get(i).isBorder()) {
                    Color color = Optional.ofNullable(titleSytle.get(i).getBorderColor()).orElse(new Color(0, 0, 0));
                    BorderStyle borderStyle = Optional.ofNullable(titleSytle.get(i).getBorderStyle()).orElse(BorderStyle.THIN);

                    setBorder(titleStyle1, borderStyle, new XSSFColor(color));
                } else {

                    setBorder(titleStyle1, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));
                }

            } else {


                titleFont.setFontName("simsun");
                titleFont.setBold(true);
                titleFont.setFontHeightInPoints((short) 9);
                titleFont.setColor(IndexedColors.BLACK.index);


//        titleStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
//        titleStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                titleStyle1.setAlignment(HorizontalAlignment.CENTER);
                titleStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
                titleStyle1.setFillForegroundColor(new XSSFColor(new Color(230, 230, 230)));
                //titleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                titleStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                titleStyle1.setFont(titleFont);
                setBorder(titleStyle1, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));
            }


            Row titleRow = sheet.createRow(rowIndex);
            // titleRow.setHeightInPoints(25);
            colIndex = 0;

            int titleLine = 0;
            //    for (int i = 0; i < titles.size(); i++) {

            List<ExclMultilieRows> rowData = titles.get(i);
            Row dataRow = sheet.createRow(rowIndex);
            // dataRow.setHeightInPoints(25);
            //  colIndex = 0;


            if (CollectionUtil.isNotBlank(rowData)) {
                for (int j = 0; j < rowData.size(); j++) {


                    // if (rowData.get(j).getValue() != null) {
                    ExclMultilieRows cellData = rowData.get(j);
                    Cell cell = dataRow.createCell(j);


                    if (cellData != null && cellData.getValue() != null) {

                        //                            if(cellData.getValue().equals("11738885")  )
                        //                                System.out.println("11738885");

                        cell.setCellValue(cellData.getValue());
                    } else {
                        cell.setCellValue("");
                    }

                    titleStyle1.setWrapText(true);
                    cell.setCellStyle(titleStyle1);
                    if (cellData.getRowSize() > 1 && cellData.getColSize() > 1) {

                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + cellData.getRowSize() - 1, j, j + cellData.getColSize() - 1));
                    } else if (cellData.getRowSize() > 1) {

                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + cellData.getRowSize() - 1, j, j));
                    } else if (cellData.getColSize() > 1) {

                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, j, j + cellData.getColSize() - 1));
                    }

                    // }
                    //  colIndex++;
                }
            }


            rowIndex++;
        }


        return rowIndex;
    }

    private static int writeRowsToExcel(XSSFWorkbook wb, Sheet sheet, List<List<Object>> rows, List<ExcelFontStyle> rowStytle, int rowIndex, List<ExcelMergedRegion> rowsMergedRegions) {
        int colIndex = 0;

        int startRow = rowIndex;

        for (int i = 0; i < rows.size(); i++) {
            List<Object> rowData = rows.get(i);
            Font dataFont = wb.createFont();

            XSSFCellStyle dataStyle = wb.createCellStyle();
            Row dataRow = sheet.createRow(rowIndex);


            if (CollectionUtil.isNotBlank(rowStytle) && i < rowStytle.size()) {
                if (Optional.ofNullable(rowStytle.get(i).getFontName()).isPresent()) {
                    dataFont.setFontName(rowStytle.get(i).getFontName());
                } else {
                    dataFont.setFontName("宋体");
                }

                dataFont.setBold(rowStytle.get(i).isBold());

                if (rowStytle.get(i).getFontHeightInPoints() != 0) {
                    dataFont.setFontHeightInPoints(rowStytle.get(i).getFontHeightInPoints());
                } else {
                    dataFont.setFontHeightInPoints((short) 9);
                }

                if (rowStytle.get(i).getColor() != 0) {
                    dataFont.setColor(rowStytle.get(i).getColor());
                } else {
                    dataFont.setColor(IndexedColors.BLACK.index);
                }

                if (Optional.ofNullable(rowStytle.get(i).getHorizontalAlignment()).isPresent()) {
                    dataStyle.setAlignment(rowStytle.get(i).getHorizontalAlignment());
                } else {
                    dataStyle.setAlignment(HorizontalAlignment.CENTER);
                }

                if (Optional.ofNullable(rowStytle.get(i).getVerticalAlignment()).isPresent()) {
                    dataStyle.setVerticalAlignment(rowStytle.get(i).getVerticalAlignment());
                } else {
                    dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                }

                if (Optional.ofNullable(rowStytle.get(i).getFillForegroundColor()).isPresent()) {
                    dataStyle.setFillForegroundColor(new XSSFColor(rowStytle.get(i).getFillForegroundColor()));
                } else {
                    dataStyle.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255)));
                }

                if (Optional.ofNullable(rowStytle.get(i).getFillPatternType()).isPresent()) {
                    dataStyle.setFillPattern(rowStytle.get(i).getFillPatternType());
                } else {
                    dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }

                if (Optional.ofNullable(rowStytle.get(i).getHeight()).isPresent()) {
                    dataRow.setHeightInPoints(rowStytle.get(i).getHeight());
                }


                dataStyle.setWrapText(rowStytle.get(i).isWrapText());

                dataStyle.setFont(dataFont);
                if (rowStytle.get(i).isBorder()) {
                    Color color = Optional.ofNullable(rowStytle.get(i).getBorderColor()).orElse(new Color(0, 0, 0));
                    BorderStyle borderStyle = Optional.ofNullable(rowStytle.get(i).getBorderStyle()).orElse(BorderStyle.THIN);

                    setBorder(dataStyle, borderStyle, new XSSFColor(color));
                } else {

                    setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));
                }

            } else {


                dataFont.setFontName("simsun");
                dataFont.setFontHeightInPoints((short) 9);
                dataFont.setColor(IndexedColors.BLACK.index);


                //   dataStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                //    dataStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                dataStyle.setAlignment(HorizontalAlignment.CENTER);
                dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);


                dataStyle.setFont(dataFont);
                setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(0, 0, 0)));


            }


            // dataRow.setHeightInPoints(25);
            colIndex = 0;


            if (CollectionUtil.isNotBlank(rowData)) {
                for (Object cellData : rowData) {
                    Cell cell = dataRow.createCell(colIndex);
                    if (cellData != null) {
                        cell.setCellValue(cellData.toString());
                    } else {
                        cell.setCellValue("");
                    }

                    dataStyle.setWrapText(true);
                    cell.setCellStyle(dataStyle);
                    colIndex++;
                }
            }
            //行列合并 区域
            if (CollectionUtil.isNotBlank(rowsMergedRegions) && i < rowsMergedRegions.size() && rowsMergedRegions.get(i).isMerged()) {
                sheet.addMergedRegion(new CellRangeAddress(rowsMergedRegions.get(i).getFirstRow() + startRow, rowsMergedRegions.get(i).getLastRow() + startRow, rowsMergedRegions.get(i).getFirstCol(), rowsMergedRegions.get(i).getLastCol()));
            }

            rowIndex++;
        }


        return rowIndex;
    }

    private static int writeMultilieRowsToExcel(XSSFWorkbook wb, Sheet sheet, List<List<ExclMultilieRows>> rows, List<ExcelFontStyle> rowStytle, int rowIndex, List<ExcelMergedRegion> mergedRegions) {
        int colIndex = 0;

        int startRow = rowIndex;

        AtomicInteger at = new AtomicInteger(0);

        for (int i = 0; i < rows.size(); i++) {
            at.set(i);
            Optional<ExcelMergedRegion> cantClinicExcelMergedRegion = Optional.empty();
            if (mergedRegions != null) {
                cantClinicExcelMergedRegion = mergedRegions.parallelStream().filter(f -> f.getFirstRow() == at.intValue() + 1 && f.getLastRow() == at.intValue() + 1).findFirst();
            }

            Optional<ExcelFontStyle> rowSty = Optional.empty();

            //查找相同行号的样式，为-99 时为默认样式
            if (rowStytle != null) {
                rowSty = rowStytle.parallelStream().filter(f -> f.getRow() == at.intValue() + 1 || f.getRow() == -99).findFirst();
            }


            Font dataFont = wb.createFont();

            XSSFCellStyle dataStyle = wb.createCellStyle();
            Row dataRow = sheet.createRow(rowIndex);

            if (rowSty.isPresent()) {
                if (rowSty.get().getFontName() != null) {
                    dataFont.setFontName(rowSty.get().getFontName());
                } else {
                    dataFont.setFontName("宋体");
                }

                dataFont.setBold(rowSty.get().isBold());

                if (rowSty.get().getFontHeightInPoints() != 0) {
                    dataFont.setFontHeightInPoints(rowSty.get().getFontHeightInPoints());
                } else {
                    dataFont.setFontHeightInPoints((short) 9);
                }

                if (rowSty.get().getColor() != 0) {
                    dataFont.setColor(rowSty.get().getColor());
                } else {
                    dataFont.setColor(IndexedColors.BLACK.index);
                }

                if (rowSty.get().getHorizontalAlignment() != null) {
                    dataStyle.setAlignment(rowSty.get().getHorizontalAlignment());
                } else {
                    dataStyle.setAlignment(HorizontalAlignment.CENTER);
                }

                if (rowSty.get().getVerticalAlignment() != null) {
                    dataStyle.setVerticalAlignment(rowSty.get().getVerticalAlignment());
                } else {
                    dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                }

                if (rowSty.get().getFillForegroundColor() != null) {
                    dataStyle.setFillForegroundColor(new XSSFColor(rowSty.get().getFillForegroundColor()));
                } else {
                    dataStyle.setFillForegroundColor(new XSSFColor(new Color(255, 255, 255)));
                }

                if (rowSty.get().getFillPatternType() != null) {
                    dataStyle.setFillPattern(rowSty.get().getFillPatternType());
                } else {
                    dataStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                }

                if (rowSty.get().getHeight() != null) {
                    dataRow.setHeightInPoints(rowSty.get().getHeight());
                }


                dataStyle.setWrapText(rowSty.get().isWrapText());

                dataStyle.setFont(dataFont);
                if (rowSty.get().isBorder()) {
                    Color color = Optional.ofNullable(rowSty.get().getBorderColor()).orElse(new Color(0, 0, 0));
                    BorderStyle borderStyle = Optional.ofNullable(rowSty.get().getBorderStyle()).orElse(BorderStyle.THIN);

                    setBorder(dataStyle, borderStyle, new XSSFColor(color));
                } else {

                    setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));
                }
            } else {


                dataFont.setFontName("simsun");
                dataFont.setFontHeightInPoints((short) 9);
                dataFont.setColor(IndexedColors.BLACK.index);


                //   dataStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
                //    dataStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
                dataStyle.setAlignment(HorizontalAlignment.CENTER);
                dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);


                dataStyle.setFont(dataFont);
                setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new Color(230, 230, 230)));


            }


            List<ExclMultilieRows> rowData = rows.get(i);

            // dataRow.setHeightInPoints(25);
            //  colIndex = 0;


            if (CollectionUtil.isNotBlank(rowData)) {
                for (int j = 0; j < rowData.size(); j++) {


                    // if (rowData.get(j).getValue() != null) {
                    ExclMultilieRows cellData = rowData.get(j);
                    Cell cell = dataRow.createCell(j);


                    if (cellData != null && cellData.getValue() != null) {

                        //                            if(cellData.getValue().equals("11738885")  )
                        //                                System.out.println("11738885");

                        cell.setCellValue(cellData.getValue());
                    } else {
                        cell.setCellValue("");
                    }


                    if (!cantClinicExcelMergedRegion.isPresent()) {
                        if (cellData.getRowSize() > 1 && cellData.getColSize() > 1) {

                            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + cellData.getRowSize() - 1, j, j + cellData.getColSize() - 1));
                        } else if (cellData.getRowSize() > 1) {

                            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + cellData.getRowSize() - 1, j, j));
                        } else if (cellData.getColSize() > 1) {

                            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, j, j + cellData.getColSize() - 1));
                        }
                    }
                    // sheet.addMergedRegion(new CellRangeAddress(startRow + i, startRow + i + cellData.getRowSize() - 1, j, j));
                    // }
                    //  colIndex++;

                    dataStyle.setWrapText(true);
                    cell.setCellStyle(dataStyle);


                }
            }


            //行列合并 区域
            if (cantClinicExcelMergedRegion.isPresent()) {
                sheet.addMergedRegion(new CellRangeAddress(cantClinicExcelMergedRegion.get().getFirstRow() + startRow, cantClinicExcelMergedRegion.get().getLastRow() + startRow, cantClinicExcelMergedRegion.get().getFirstCol(), cantClinicExcelMergedRegion.get().getLastCol()));
            }


            rowIndex++;
        }


        return rowIndex;
    }

    private static void autoSizeColumns(Sheet sheet, int columnNumber) {

        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);

            sheet.autoSizeColumn(i, true);
            int newWidth = (int) (sheet.getColumnWidth(i) + 100);
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, newWidth);
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }

    private static void setSizeColumns(Sheet sheet, List<Integer> colsWidth) {

        for (int i = 0; i < colsWidth.size(); i++) {

            if (colsWidth.get(i) > 0) {
                sheet.setColumnWidth(i, colsWidth.get(i));
            }

        }


    }

    private static void setBorder(XSSFCellStyle style, BorderStyle border, XSSFColor color) {
        style.setBorderTop(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
        style.setBorderBottom(border);
        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, color);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, color);
    }

    /**
     * 合并单元格
     *
     * @param sheet
     * @param cellLine
     * @param startRow
     * @param endRow
     * @param workbook
     * @param mergeBasis
     */
    private static void mergedRegion2003(HSSFSheet sheet, int cellLine, int startRow, int endRow, HSSFWorkbook workbook, Integer[] mergeBasis) {
        HSSFCellStyle style = workbook.createCellStyle();           // 样式对象
        style.setVerticalAlignment(VerticalAlignment.CENTER);  // 垂直
        style.setAlignment(HorizontalAlignment.CENTER);             // 水平
        String s_will = sheet.getRow(startRow).getCell(cellLine).getStringCellValue();  // 获取第一行的数据,以便后面进行比较
        int count = 0;
        Set<Integer> set = new HashSet<>();
        CollectionUtils.addAll(set, mergeBasis);
        for (int i = startRow + 1; i <= endRow; i++) {
            String s_current = sheet.getRow(i).getCell(cellLine).getStringCellValue();
            if (s_will.equals(s_current)) {
                boolean isMerge = true;
                if (!set.contains(cellLine)) {//如果不是作为基准列的列 需要所有基准列都相同
                    for (int j = 0; j < mergeBasis.length; j++) {
                        if (!sheet.getRow(i).getCell(mergeBasis[j]).getStringCellValue().equals(sheet.getRow(i - 1).getCell(mergeBasis[j]).getStringCellValue())) {
                            isMerge = false;
                        }
                    }
                } else {//如果作为基准列的列 只需要比较列号比本列号小的列相同
                    for (int j = 0; j < mergeBasis.length && mergeBasis[j] < cellLine; j++) {
                        if (!sheet.getRow(i).getCell(mergeBasis[j]).getStringCellValue().equals(sheet.getRow(i - 1).getCell(mergeBasis[j]).getStringCellValue())) {
                            isMerge = false;
                        }
                    }
                }
                if (isMerge) {
                    count++;
                } else {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, startRow + count, cellLine, cellLine));
                    startRow = i;
                    s_will = s_current;
                    count = 0;
                }
            } else {
                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow + count, cellLine, cellLine));
                startRow = i;
                s_will = s_current;
                count = 0;
            }
            if (i == endRow && count > 0) {
                sheet.addMergedRegion(new CellRangeAddress(startRow, startRow + count, cellLine, cellLine));
            }
        }
    }

    /**
     * 判断Excel的版本,获取Workbook
     *
     * @param in
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(InputStream in, File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(EXCEL_XLS)) {  //Excel 2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(EXCEL_XLSX)) {  // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    /**
     * 判断Word的版本,
     *
     * @param in
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWork(InputStream in, File file) throws IOException {
        Workbook wb = null;
        if (file.getName().endsWith(WORD_DOC)) {  //Excel 2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith(WORD_DOCX)) {  // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    /**
     * 判断文件是否是excel
     *
     * @throws Exception
     */
    public static void checkExcelVaild(File file) throws Exception {
        if (!file.exists()) {
            throw new Exception("文件不存在");
        }
        if (!(file.isFile() && (file.getName().endsWith(EXCEL_XLS) || file.getName().endsWith(EXCEL_XLSX)))) {
            throw new Exception("文件不是Excel");
        }
    }

    private static Object getValue(Cell cell) {
        Object obj = null;
        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                obj = cell.getBooleanCellValue();
                break;
            case ERROR:
                obj = cell.getErrorCellValue();
                break;
            case NUMERIC:
                obj = cell.getNumericCellValue();
                break;
            case STRING:
                obj = cell.getStringCellValue();
                break;
            default:
                break;
        }
        return obj;
    }

    public static List readExcel(File file) {
        List reList = new Vector();

        try (FileInputStream in = new FileInputStream(file);) // 文件流
        {

            checkExcelVaild(file);
            Workbook workbook = getWorkbok(in, file);
            //Workbook workbook = WorkbookFactory.create(is); // 这种方式 Excel2003/2007/2010都是可以处理的

            int sheetCount = workbook.getNumberOfSheets();
            Sheet sheet = workbook.getSheetAt(0);   // 遍历第一个Sheet

            // 为跳过第一行目录设置count
            // 跳过第一和第二行的目录
//            if(count < 2 ) {
//                count++;
//                continue;
//            }
            int count = 0;
            for (Row row : sheet) {
                //Map map=new HashMap();
                List cellList = new Vector();
                //如果当前行没有数据，跳出循环
//                if(row.getCell(0).toString().equals("")){
//                    break;
//                }
                //获取总列数(空格的不计算)
                int columnTotalNum = row.getPhysicalNumberOfCells();

                System.out.println("总列数：" + columnTotalNum);

                System.out.println("最大列数：" + row.getLastCellNum());

                int end = row.getLastCellNum();
                for (int i = 0; i < end; i++) {
                    Cell cell = row.getCell(i);
                    if (cell == null) {
                        System.out.print("null" + "\t");
                        continue;
                    }

                    Object obj = getValue(cell);

                    // System.out.print(obj + "\t");
                    if (obj != null) {
                        cellList.add(obj);
                    }
                }

                reList.add(cellList);
            }

            return reList;
        } catch (Exception e) {

            e.printStackTrace();
        }


        return null;

    }

    /**
     * 替换Excel模板文件内容
     *
     * @param item           文档数据
     * @param sourceFilePath Excel模板文件路径
     * @param targetFilePath Excel生成文件路径
     */
    public static boolean replaceModel(Map item, String sourceFilePath, String targetFilePath) {
        try {
            File file = ResourceUtils.getFile(sourceFilePath);
            if (file.getName().endsWith(EXCEL_XLS)) {  //Excel 2003
                return replaceModel2003(item, file, targetFilePath);
            } else if (file.getName().endsWith(EXCEL_XLSX)) {  // Excel 2007/2010
                return replaceModel2017(item, file, targetFilePath);
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        return false;
    }

    public static boolean replaceModel(Map item, String sourceFilePath, HttpServletResponse response) {
        try {
            File file = ResourceUtils.getFile(sourceFilePath);
            if (file.getName().endsWith(EXCEL_XLS)) {  //Excel 2003
                return replaceModel2003(item, file, response.getOutputStream());
            } else if (file.getName().endsWith(EXCEL_XLSX)) {  // Excel 2007/2010
                return replaceModel2017(item, file, response.getOutputStream());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;
    }

    public static boolean replaceWordModel(List<Map> item, String sourceFilePath, HttpServletResponse response) {
        try (InputStream fileInputStream = new FileInputStream(sourceFilePath)) {


            //首先定义一个XWPFDocument 集合 这个对象可以进行word 解析 合并 还有下载都离不开这个对象
            if (CollectionUtil.isBlank(item)) {
                return false;
            }

            //File file = ResourceUtils.getFile(sourceFilePath);
            if (item.size() == 1) {
                return replaceWordModel(item, fileInputStream, response);
            } else {
                String date = DateUtil.toString(DateUtil.getDateline(), DateUtil.DATE_FMTHM);

                String[] srcDocxs = new String[item.size()];
                for (int i = 0; i < item.size(); i++) {
                    //多个文档

                    String destinationNameX = "cfqian" + date + "-" + new Random().nextInt(100000) + "-tmp.docx";

                    String fileStringX = copyReplExcel(sourceFilePath, destinationNameX, item.get(i));
                    if (StringUtil.isBlank(fileStringX)) {
                        break;
                    }


                    srcDocxs[i] = fileStringX;
                }

                String destDocx = System.getProperty("user.dir") + "/excltmp" + "/new-cfqian" + date + "-" + new Random().nextInt(100000) + "-tmp.docx";
                mergeDoc(srcDocxs, destDocx, item);
                try (InputStream fileInputStreamX = new FileInputStream(destDocx)) {


                    XWPFDocument document = new XWPFDocument(fileInputStreamX);


                    document.write(response.getOutputStream());
                    File fileX = ResourceUtils.getFile(destDocx);
                    fileX.delete();
                } catch (Exception ex) {

                    ex.printStackTrace();
                }
                return true;

            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return false;

    }

    /**
     * 合并docx文件
     *
     * @param srcDocxs 需要合并的目标docx文件
     * @param destDocx 合并后的docx输出文件
     */
    public static void mergeDoc(String[] srcDocxs, String destDocx, List<Map> items) {
        int index = 0;
        OutputStream dest = null;
        List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
        int length = null == srcDocxs ? 0 : srcDocxs.length;
        /**
         * 循环获取每个docx文件的OPCPackage对象
         */
        for (int i = 0; i < length; i++) {
            String doc = srcDocxs[i];
            OPCPackage srcPackage = null;
            try {
                srcPackage = OPCPackage.open(doc);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (null != srcPackage) {
                opcpList.add(srcPackage);
            }
        }

        int opcpSize = opcpList.size();
        //获取的OPCPackage对象大于0时，执行合并操作
        if (opcpSize > 0) {
            try {
                dest = new FileOutputStream(destDocx);


                XWPFDocument src1Document = new XWPFDocument(opcpList.get(0));
                CTBody src1Body = src1Document.getDocument().getBody();


                //OPCPackage大于1的部分执行合并操作
                if (opcpSize > 1) {
                    for (int i = 1; i < opcpSize; i++) {
                        index = i;
                        OPCPackage src2Package = opcpList.get(i);
                        XWPFDocument src2Document = new XWPFDocument(src2Package);
                        CTBody src2Body = src2Document.getDocument().getBody();
                        appendBody(src1Body, src2Body, items);
                    }
                }
                //将合并的文档写入目标文件中
                src1Document.write(dest);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                //注释掉以下部分，去除影响目标文件srcDocxs。
				/*for (OPCPackage opcPackage : opcpList) {
					if(null != opcPackage){
						try {
							opcPackage.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}*/
                //关闭流
                IOUtils.closeQuietly(dest);
            }
        }


    }

    /**
     * 合并文档内容
     *
     * @param src    目标文档
     * @param append 要合并的文档
     * @throws Exception
     */
    private static void appendBody(CTBody src, CTBody append, List<Map> items) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();

        String appendString = append.xmlText(optionsOuter);


        String srcString = src.xmlText();

        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart = srcString.substring(srcString.indexOf(">") + 1,
                srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart = appendString.substring(appendString.indexOf(">") + 1,
                appendString.lastIndexOf("<"));
        CTBody makeBody = CTBody.Factory.parse(prefix + mainPart + addPart
                + sufix);
        src.set(makeBody);
    }

    public static boolean replaceWordModel(List<Map> item, InputStream fileInputStream, HttpServletResponse response) {

        boolean bool = true;
        try (
                HWPFDocument doc = new HWPFDocument(fileInputStream);
                // XWPFDocument doc = new XWPFDocument(fileInputStream);


        ) {

            ByteArrayOutputStream ostream = new ByteArrayOutputStream();
            ServletOutputStream servletOS = response.getOutputStream();


            for (int i = 0; i < item.size(); i++) {
                Range range = doc.getRange();


                Set<String> keySet = item.get(i).keySet();
                Iterator<String> it = keySet.iterator();
                while (it.hasNext()) {
                    String text = it.next();

                    range.replaceText(text, String.valueOf(item.get(i).get(text)));
                }

                doc.write(ostream);
                servletOS.write(ostream.toByteArray());
                servletOS.flush();

            }


            servletOS.close();
            // 输出文件
            //   FileOutputStream fileOut = new FileOutputStream(targetFilePath);

            //  fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;

    }

    public static boolean replaceModel2017(Map item, File file, String targetFilePath) {


        boolean bool = true;
        try (
                //POIFSFileSystem fs  =new POIFSFileSystem(new FileInputStream(file));
                XSSFWorkbook wb = new XSSFWorkbook(file);) {

            XSSFCellStyle cellStyle = wb.createCellStyle();

            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                if (row != null) {
                    int num = row.getLastCellNum();
                    for (int i = 0; i < num; i++) {
                        XSSFCell cell = row.getCell(i);
                        if (cell != null) {
                            cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                        }
                        if (cell == null || cell.getStringCellValue() == null) {
                            continue;
                        }
                        String value = cell.getStringCellValue();
                        if (!"".equals(value)) {
                            Set<String> keySet = item.keySet();
                            Iterator<String> it = keySet.iterator();
                            while (it.hasNext()) {
                                String text = it.next();
                                if (value.equalsIgnoreCase(text)) {

                                    if (text.equalsIgnoreCase("\n")) {
                                        cellStyle = cell.getCellStyle();
                                        cellStyle.setWrapText(true);
                                        cell.setCellStyle(cellStyle);
                                    }
//                                    cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
//                                    cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
//                                    cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
//                                    cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框
                                    cell.setCellValue(new XSSFRichTextString(String.valueOf(item.get(text))));

                                    break;
                                }
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }

            // 输出文件
            FileOutputStream fileOut = new FileOutputStream(targetFilePath);
            wb.write(fileOut);
            fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;

    }

    public static boolean replaceModel2017(Map item, File file, OutputStream out) {

        boolean bool = true;
        try (
                //POIFSFileSystem fs  =new POIFSFileSystem(new FileInputStream(file));
                XSSFWorkbook wb = new XSSFWorkbook(file);) {

            XSSFCellStyle cellStyle = (XSSFCellStyle) wb.createCellStyle();


            XSSFSheet sheet = wb.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                XSSFRow row = (XSSFRow) rows.next();
                if (row != null) {
                    int num = row.getLastCellNum();
                    for (int i = 0; i < num; i++) {
                        XSSFCell cell = row.getCell(i);
                        if (cell != null) {
                            cell.setCellType(XSSFCell.CELL_TYPE_STRING);

                        }
                        if (cell == null || cell.getStringCellValue() == null) {
                            continue;
                        }
                        String value = cell.getStringCellValue();
                        if (!"".equals(value)) {
                            Set<String> keySet = item.keySet();
                            Iterator<String> it = keySet.iterator();
                            while (it.hasNext()) {
                                String text = it.next();
                                if (value.toLowerCase().contains(text.toLowerCase())) {

                                    // cell = row.createCell((short) 0);

                                    if (String.valueOf(item.get(text) == null ? "" : item.get(text)).contains("\n")) {
                                        cellStyle = cell.getCellStyle();
                                        cellStyle.setWrapText(true);
                                        cell.setCellStyle(cellStyle);
                                    }

//                                    cellStyle.setBorderBottom(XSSFCellStyle.BORDER_THIN); //下边框
//                                    cellStyle.setBorderLeft(XSSFCellStyle.BORDER_THIN);//左边框
//                                    cellStyle.setBorderTop(XSSFCellStyle.BORDER_THIN);//上边框
//                                    cellStyle.setBorderRight(XSSFCellStyle.BORDER_THIN);//右边框

                                    if (item.get(text) != null) {
                                        cell.setCellValue(cell.getStringCellValue().replaceAll(text, String.valueOf(item.get(text))));
                                    } else {
                                        cell.setCellValue(cell.getStringCellValue().replaceAll(text, ""));
                                    }

                                }
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }

            // 输出文件
            //   FileOutputStream fileOut = new FileOutputStream(targetFilePath);
            wb.write(out);
            //  fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;

    }

    public static boolean replaceModel2003(Map item, File file, String targetFilePath) {

        boolean bool = true;
        try (POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file)); HSSFWorkbook wb = new HSSFWorkbook(fs);) {

            HSSFCellStyle cellStyle = wb.createCellStyle();

            HSSFSheet sheet = wb.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                HSSFRow row = (HSSFRow) rows.next();
                if (row != null) {
                    int num = row.getLastCellNum();
                    for (int i = 0; i < num; i++) {
                        HSSFCell cell = row.getCell(i);
                        if (cell != null) {
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        }
                        if (cell == null || cell.getStringCellValue() == null) {
                            continue;
                        }
                        String value = cell.getStringCellValue();
                        if (!"".equals(value)) {
                            Set<String> keySet = item.keySet();
                            Iterator<String> it = keySet.iterator();
                            while (it.hasNext()) {
                                String text = it.next();
                                if (value.equalsIgnoreCase(text)) {

                                    if (text.equalsIgnoreCase("\n")) {
                                        cellStyle = cell.getCellStyle();
                                        cellStyle.setWrapText(true);
                                        cell.setCellStyle(cellStyle);

                                    }
//                                    cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
//                                    cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
//                                    cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
//                                    cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
                                    cell.setCellValue(new HSSFRichTextString(String.valueOf(item.get(text))));

                                    break;
                                }
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }

            // 输出文件
            FileOutputStream fileOut = new FileOutputStream(targetFilePath);
            wb.write(fileOut);
            fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;

    }

    public static boolean replaceModel2003(Map item, File file, OutputStream out) {

        boolean bool = true;
        try (POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file)); HSSFWorkbook wb = new HSSFWorkbook(fs);) {
            HSSFCellStyle cellStyle = wb.createCellStyle();

            HSSFSheet sheet = wb.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            while (rows.hasNext()) {
                HSSFRow row = (HSSFRow) rows.next();
                if (row != null) {
                    int num = row.getLastCellNum();
                    for (int i = 0; i < num; i++) {
                        HSSFCell cell = row.getCell(i);
                        if (cell != null) {
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                        }
                        if (cell == null || cell.getStringCellValue() == null) {
                            continue;
                        }
                        String value = cell.getStringCellValue();
                        if (!"".equals(value)) {
                            Set<String> keySet = item.keySet();
                            Iterator<String> it = keySet.iterator();
                            while (it.hasNext()) {
                                String text = it.next();
                                if (value.equalsIgnoreCase(text)) {

                                    if (text.equalsIgnoreCase("\n")) {
                                        cellStyle = cell.getCellStyle();
                                        cellStyle.setWrapText(true);
                                        cell.setCellStyle(cellStyle);

                                    }
//                                    cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
//                                    cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
//                                    cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
//                                    cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

                                    cell.setCellValue(new HSSFRichTextString(String.valueOf(item.get(text))));

                                    break;
                                }
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }

            // 输出文件
            //   FileOutputStream fileOut = new FileOutputStream(targetFilePath);
            wb.write(out);
            //   fileOut.close();

        } catch (Exception e) {
            bool = false;
            e.printStackTrace();
        }
        return bool;

    }

    /**
     * 拷贝excl
     *
     * @param sourcePath      带路径源文件
     * @param destinationName 不带路径目标文件名字
     * @return
     */
    public static String copyExcel(String sourcePath, String destinationName) {
        String reString = "";

        String excelWork = SystemConfigure.getExcelTemp();

        File directory = new File(excelWork);

        if (excelWork == null || !directory.exists()) {
            excelWork = System.getProperty("user.dir") + "/excltmp";
            directory = new File(excelWork);
            if (!directory.exists()) {
                directory.mkdir();
            }
        }


        String excelFile = excelWork + "/" + destinationName;
        int byteread = 0; // 读取的字节数
        try {
            File file = ResourceUtils.getFile(sourcePath);
            try (

                    InputStream in = new FileInputStream(file); OutputStream out = new FileOutputStream(excelFile);) {

                byte[] buffer = new byte[1024];

                while ((byteread = in.read(buffer)) != -1) {
                    out.write(buffer, 0, byteread);
                }
                return reString = excelFile;

            } catch (Exception ex) {
                SystemLogUtil.error("excl文件拷贝失败！！" + ex.toString());
                return null;
            }

        } catch (Exception e) {

            SystemLogUtil.error("文件路径不对！！" + e.toString());
            return null;

        }


        //return reString;
    }

    public static String copyReplExcel(String sourcePath, String destinationName, Map<String, String> items) {
        // String reString = "";

        String excelWork = SystemConfigure.getExcelTemp();

        File directory = new File(excelWork);

        if (excelWork == null || !directory.exists()) {
            excelWork = System.getProperty("user.dir") + "/excltmp";
            directory = new File(excelWork);
            if (!directory.exists()) {
                directory.mkdir();
            }
        }


        String excelFile = excelWork + "/" + destinationName;
        try {
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(sourcePath));
            // 替换段落中的指定文字
            Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
            while (itPara.hasNext()) {
                XWPFParagraph paragraph = itPara.next();
                String oneparaString = paragraph.getText();

                if (StringUtil.isBlank(oneparaString)) {
                    continue;
                }

                for (Map.Entry<String, String> entry : items.entrySet()) {
                    if (Strings.isBlank(entry.getKey()) || entry.getValue() == null) {
                        continue;
                    }
                    oneparaString = oneparaString.replace(entry.getKey(), entry.getValue());
                }

            }


            FileOutputStream outStream = new FileOutputStream(destinationName);
            document.write(outStream);
            outStream.close();

            return destinationName;


        } catch (Exception e) {

            SystemLogUtil.error("文件路径不对！！" + e.toString());

        }


        return null;
    }
/***************************************************************************************************导入******************************************************************************/

    /**
     * 得到Excel表中的值
     * * * @param hssfCell
     * * Excel中的每一个格子
     * * @return Excel中每一个格子中的值
     */
    @SuppressWarnings("static-access")
    public static String getValue(HSSFCell hssfCell) {
        if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
            // 返回布尔类型的值
            return String.valueOf(hssfCell.getBooleanCellValue());
        } else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
            // 返回数值类型的值
            return String.valueOf(hssfCell.getNumericCellValue());
        } else {
            // 返回字符串类型的值
            return String.valueOf(hssfCell.getStringCellValue());
        }
    }

    //判断Excel倒入数据类型，转换为数据库可识别的数据类型
    @SuppressWarnings({"static-access", "unused"})
    public static String getCellTypes(Cell cell) {
        String cellValue = null;
        if (null != cell) {
            // 以下是判断数据的类型
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_NUMERIC: // 数字 // 处理日期格式、时间格式
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date d = cell.getDateCellValue();
                        DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = formater.format(d);
                    } else {
                        cellValue = cell.getNumericCellValue() + "";
                    }
                    break;
                case HSSFCell.CELL_TYPE_STRING: // 字符串
                    cellValue = cell.getStringCellValue();
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN: //  BOOL
                    cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                    break;
                case HSSFCell.CELL_TYPE_FORMULA: // 公式 //
                    cellValue = cell.getCellFormula() + "";
//                    try {
//                        DecimalFormat df = new DecimalFormat("0.0000");
//                        cellValue = String.valueOf(df.format(cell.getNumericCellValue()));
//                    } catch (IllegalStateException e) {
//                        cellValue = String.valueOf(cell.getRichStringCellValue());
//                    }
                    break;
                case HSSFCell.CELL_TYPE_BLANK: // 空值
                    cellValue = "";
                    break;
                case HSSFCell.CELL_TYPE_ERROR: // 故障
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
        }
        return cellValue;
    }


    /**
     * 描述：验证EXCEL文件
     * * @param filePath
     * * @return
     */
    public static ExcelValidate validateExcel(String filePath) {
        ExcelValidate validateMsg = new ExcelValidate();
        /** 检查文件是否存在 */
        File file = new File(filePath);

        boolean a = file.exists();
        if (file == null || !file.exists()) {
            validateMsg.setMsg("文件不存在");
            validateMsg.setState(false);
            return validateMsg;
        }
        if (filePath == null || !(file.getName().endsWith(EXCEL_XLSX) || file.getName().endsWith(EXCEL_XLS))) {
            validateMsg.setMsg("文件名不是excel格式");
            validateMsg.setState(false);
            return validateMsg;
        }
        validateMsg.setMsg("成功");
        validateMsg.setState(true);
        return validateMsg;
    }

    /**
     * @param filePath
     * @return
     * @描述：是否是2003的excel，返回true是2003
     */
    public static boolean isExcel2003(String filePath) {
        return filePath.matches("^.+\\.(?i)(xls)$");
    }

    /**
     * @param filePath
     * @return
     * @描述：是否是2007的excel，返回true是2007
     */
    public static boolean isExcel2007(String filePath) {
        return filePath.matches("^.+\\.(?i)(xlsx)$");
    }

    /**
     * 描述 :读EXCEL文件
     *
     * @param Mfile
     * @param claszz
     * @return
     */
    public static ExcelValidate getExcelInfo(MultipartFile Mfile, Class claszz) {

        ExcelValidate excelValidate = new ExcelValidate(false, "错误！");
        //把spring文件上传的MultipartFile转换成File
        String fileName = Mfile.getOriginalFilename();

        //根据新建的文件实例化输入流
        InputStream is = null;
        Workbook wb = null;
        try {
            is = Mfile.getInputStream();
            //根据excel里面的内容读取客户信息
            if (ExcelUtils.isExcel2003(fileName)) {
                //当excel是2003时
                wb = new HSSFWorkbook(is);
            } else if (ExcelUtils.isExcel2007(fileName)) {
                //当excel是2007时
                wb = new XSSFWorkbook(is);
            }

            excelValidate = readExcelValue(wb, claszz);


            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    is = null;
                    e.printStackTrace();
                }
            }
        }
        return excelValidate;
    }


    /**
     * 读取Excel里面的信息
     * * @param wb
     * * @return
     */
    private static ExcelValidate readExcelValue(Workbook wb, Class claszz) {
        ExcelValidate excelValidate = new ExcelValidate(false, "错误！");
        try {


            Field[] fields = claszz.getDeclaredFields();
            //得到第一个shell
            Sheet sheet = wb.getSheetAt(0);
            //得到Excel的行数
            int totalRows = sheet.getPhysicalNumberOfRows();
            int totalCells = 0;
            //得到Excel的列数(前提是有行数)
            if (totalRows >= 1 && sheet.getRow(0) != null) {
                totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
            }
            if (fields.length != totalCells) {
                excelValidate.setState(false);
                excelValidate.setMsg("列数不匹配！！类" + claszz.getName() + "列数:" + fields.length + ",导入列数:" + totalCells);
                return excelValidate;

            }

//            Row fisrtrow = sheet.getRow(0);
//            for (int i = 0; i < totalCells; i++) {
//                Cell cell = fisrtrow.getCell(i);
//                String cellValue = getCellTypes(cell);
//                if (!fields[i].getName().equals(cellValue)) {
//
//                    excelValidate.setState(false);
//                    excelValidate.setMsg("列名:" + cellValue + "不匹配！！");
//                    return excelValidate;
//                }
//
//            }


            List userList = new ArrayList<>();

            for (int r = 1; r < totalRows; r++) {
                Row row = sheet.getRow(r);
                Object obj = claszz.newInstance();
                if (row == null) {
                    continue;
                } //
                List rowLst = new ArrayList();

                /** 循环Excel的列 */
                for (int c = 0; c < totalCells; c++) {
                    row.getCell(c).setCellType(CellType.STRING);
                    Cell cell = row.getCell(c);
                    String cellValue = getCellTypes(cell);
                    PropertyDescriptor pd = new PropertyDescriptor(fields[c].getName(), claszz);
                    Method setMethod = pd.getWriteMethod();//获得写方法
                    Method getMethod = pd.getReadMethod();//获得读方法
                    Object value = getMethod.invoke(obj);
                    if (setMethod != null) {
                        // System.out.println(object+"的字段是:"+fields[c].getName()+"，参数类型是："+fields[c].getType()+"，set的值是： "+cellValue);
                        //这里注意实体类中set方法中的参数类型，如果不是String类型则进行相对应的转换


                        if (value instanceof Double) {
                            setMethod.invoke(obj, Double.parseDouble(cellValue));
                        } else if (value instanceof Date) {
                            setMethod.invoke(obj, DateUtils.parseDate(cellValue, new String[]{"yyyy", "MM", "dd"}));
                        } else if (value instanceof String) {
                            setMethod.invoke(obj, cellValue);
                        } else if (value instanceof Boolean) {
                            setMethod.invoke(obj, Boolean.parseBoolean(cellValue));
                        } else if (value instanceof Long) {
                            setMethod.invoke(obj, Long.parseLong(cellValue));
                        } else if (value instanceof Integer) {
                            setMethod.invoke(obj, Integer.parseInt(cellValue));
                        } else if (value instanceof Float) {
                            setMethod.invoke(obj, Float.parseFloat(cellValue));
                        } else if (value instanceof Short) {
                            setMethod.invoke(obj, Short.parseShort(cellValue));
                        } else {
                            setMethod.invoke(obj, cellValue);
                        }


                        //invoke是执行set方法
                    }


                    userList.add(obj);
                }
            }
            excelValidate.setState(true);
            excelValidate.setMsg("导入成功！");
            excelValidate.setArrlist(userList);

            return excelValidate;
        } catch (SecurityException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (IntrospectionException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return excelValidate;
    }
}







