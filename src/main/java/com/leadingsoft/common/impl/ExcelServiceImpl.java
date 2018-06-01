package com.leadingsoft.common.impl;

import com.leadingsoft.common.ExcelService;
import com.leadingsoft.common.IDataCallBack;
import com.leadingsoft.common.IExcelCallBack;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

/**
 * excel导出组件
 */
public class ExcelServiceImpl implements ExcelService {

    private static final String SONG_TI_FONT = "宋体";

    @Override
    public HSSFWorkbook createWorkBook(String fileName, String[] displayNames, String[] keys, List datas, Class _class, IExcelCallBack callBack, IDataCallBack dataCallback) throws IllegalAccessException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        Font contentFont = workbook.createFont();
        contentFont.setBold(false);
        // 计算该数据表的列数
        for (int i = 0; i < displayNames.length; i++) {
            if (i == 0) {
                setColumnsWidthStyle(sheet, i, 3000);
            } else {
                setColumnsWidthStyle(sheet, i, 6000);
            }
        }
        // 设置标题行
        setHeadLine(workbook, displayNames, 0, (short) 800);

        CellStyle style = createTrueBorderCellStyle(workbook,
                HSSFColor.WHITE.index, HSSFColor.WHITE.index,
                HorizontalAlignment.CENTER, contentFont);
        style.setWrapText(true);
        Field[] fs = _class.getDeclaredFields();
        short rowNumber = 0;
        for (Object object : datas) {
            rowNumber++;
            HSSFRow row = sheet.createRow((short) rowNumber);
            row.setHeight((short) 800);
            for (int i = 0; i < keys.length; i++) {
                for (Field field : fs) {
                    field.setAccessible(true); // 设置些属性是可以访问的
                    if (keys[i].equals(field.getName())) {
                        HSSFCell cell = row.createCell(i);
                        cell.setCellType(CellType.STRING);
                        String key = keys[i];
                        String val = field.get(object) == null ? "" : field.get(object).toString();
                        val = dataCallback == null ? val : dataCallback.run(key, val);
                        cell.setCellValue(new HSSFRichTextString(val));
                        cell.setCellStyle(style);
                    }

                }

            }
        }
        if (callBack != null) {
            callBack.run(sheet);
        }

        return workbook;
    }

    @Override
    public HSSFWorkbook createWorkBook(String fileName, String[] displayNames, String[] keys, List<Map<String, Object>> datas, IExcelCallBack callBack, IDataCallBack dataCallback) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        Font contentFont = workbook.createFont();
        contentFont.setBold(false);
        // 计算该数据表的列数
        for (int i = 0; i < displayNames.length; i++) {
            if (i == 0) {
                setColumnsWidthStyle(sheet, i, 3000);
            } else {
                setColumnsWidthStyle(sheet, i, 5000);
            }
        }
        // 设置标题行
        setHeadLine(workbook, displayNames, 0, (short) 800);
        CellStyle style = createTrueBorderCellStyle(workbook,
                HSSFColor.WHITE.index, HSSFColor.WHITE.index,
                HorizontalAlignment.CENTER, contentFont);
        style.setWrapText(true);

        short rowNumber = 0;
        for (Map<String, Object> map : datas) {
            rowNumber++;
            HSSFRow row = sheet.createRow((short) rowNumber);
            row.setHeight((short) 800);
            for (int i = 0; i < keys.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellType(CellType.STRING);
                String key = keys[i].toString();
                String val = map.get(key) == null ? "" : map.get(key).toString();
                val = dataCallback == null ? val : dataCallback.run(key, val);
                cell.setCellValue(new HSSFRichTextString(val));
                cell.setCellStyle(style);
            }
        }

        if (callBack != null) {
            callBack.run(sheet);
        }
        return workbook;
    }

    @Override
    public HSSFWorkbook createWorkBook(String fileName, List<String[]> titles, String[] keys, List<Map<String, Object>> datas, IExcelCallBack callBack, IDataCallBack dataCallback) {

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        Font contentFont = workbook.createFont();
        contentFont.setBold(true);
        //处理表头
        if (titles != null && titles.size() > 0) {
            // 计算该数据表的列数
            for (int i = 0; i < titles.get(0).length; i++) {
                if (i == 0) {
                    setColumnsWidthStyle(sheet, i, 3000);
                } else {
                    setColumnsWidthStyle(sheet, i, 5000);
                }
            }
            for (int i = 0; i < titles.size(); i++) {
                if (i == 0) {
                    String[] displayNames = titles.get(i);
                    // 设置标题行
                    setHeadLine(workbook, displayNames, i, (short) 900);
                } else {
                    String[] displayNames = titles.get(i);
                    // 设置标题行
                    setHeadLine(workbook, displayNames, i, (short) 800);
                }
            }
        }

        //处理数据内容
        CellStyle style = createTrueBorderCellStyle(workbook,
                HSSFColor.WHITE.index, HSSFColor.WHITE.index,
                HorizontalAlignment.CENTER, contentFont);
        style.setWrapText(true);

        short rowNumber = 0;
        if (CollectionUtils.isNotEmpty(titles)) {
            rowNumber = (short) titles.size();
        }
        for (Map<String, Object> map : datas) {
            HSSFRow row = sheet.createRow(rowNumber++);
            row.setHeight((short) 800);

            for (int i = 0; i < keys.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellType(CellType.STRING);
                String key = keys[i];
                String val = map.get(key) == null ? "" : map.get(key).toString();
                val = dataCallback == null ? val : dataCallback.run(key, val);
                cell.setCellValue(new HSSFRichTextString(val));
                cell.setCellStyle(style);
            }
        }

        if (callBack != null) {
            callBack.run(sheet);
        }

        return workbook;
    }

    /**
     * 获取一个默认的单元格样式
     *
     * @param wb
     * @param backgroundColor
     * @param foregroundColor
     * @param halign
     * @param borderFont
     * @return
     */
    private CellStyle createTrueBorderCellStyle(HSSFWorkbook wb,
                                                short backgroundColor, short foregroundColor, HorizontalAlignment halign,
                                                Font borderFont) {
        CellStyle cs = wb.createCellStyle();
        cs.setAlignment(halign);
        cs.setVerticalAlignment(VerticalAlignment.CENTER);
        cs.setFillBackgroundColor(backgroundColor);
        cs.setFillForegroundColor(foregroundColor);
        cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cs.setFont(borderFont);
        cs.setBorderLeft(BorderStyle.THIN);
        cs.setBorderRight(BorderStyle.THIN);
        cs.setBorderTop(BorderStyle.THIN);
        cs.setBorderBottom(BorderStyle.THIN);
        return cs;
    }


    /**
     * 设置某一列的宽度
     *
     * @param sheet
     * @param width
     * @return
     */
    private HSSFSheet setColumnsWidthStyle(HSSFSheet sheet, int column, int width) {
        sheet.setColumnWidth(column, width);
        return sheet;
    }


    /**
     * 设置标题行
     *
     * @param workbook   工作表
     * @param headString 标题名称
     * @param lineNumber 行号
     * @param rowHeight  行高
     */
    private void setHeadLine(HSSFWorkbook workbook, String[] headString, int lineNumber, short rowHeight) {
        HSSFFont headFont = workbook.createFont();
        headFont.setFontName(SONG_TI_FONT);
        headFont.setBold(true);
        CellStyle style3 = createTrueBorderCellStyle(workbook,
                HSSFColor.GREY_25_PERCENT.index,
                HSSFColor.GREY_25_PERCENT.index, HorizontalAlignment.CENTER,
                headFont);
        HSSFRow row = workbook.getSheetAt(0).createRow(lineNumber);
        row.setHeight(rowHeight);
        for (int i = 0; i < headString.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style3);
            HSSFRichTextString text2 = new HSSFRichTextString(headString[i]);
            cell.setCellValue(text2);
        }
    }

}
