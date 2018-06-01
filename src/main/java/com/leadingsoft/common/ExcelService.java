package com.leadingsoft.common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.List;
import java.util.Map;

public interface ExcelService {


    /**
     * 生成Excel
     *
     * @param sheetName    表格名字
     * @param displayNames Excel字段,与keys按顺序一一对应
     * @param keys         对应数据源对象的属性
     * @param datas        数据源
     * @param _class       数据源中的对象的class
     * @param callBack     回调，对excel进一步处理,可为null
     * @param dataCallback 回调，对数据进行处理,可为null
     * @return
     */
    HSSFWorkbook createWorkBook(String sheetName, String[] displayNames,
                                String[] keys, List datas, Class _class,
                                IExcelCallBack callBack, IDataCallBack dataCallback) throws IllegalAccessException;

    /**
     * 生成Excel
     *
     * @param sheetName    表格名字
     * @param displayNames Excel字段,与keys按顺序一一对应
     * @param keys         对应数据源对象的属性
     * @param datas        数据源
     * @param callBack     回调，对excel进一步处理,可为null
     * @param dataCallback 回调，对数据进行处理,可为null
     * @return
     */
    HSSFWorkbook createWorkBook(String sheetName, String[] displayNames,
                                String[] keys, List<Map<String, Object>> datas,
                                IExcelCallBack callBack, IDataCallBack dataCallback);

    /**
     * 生成Excel,支持表头多行
     *
     * @param sheetName                   表格名字
     * @param titles                      表头 可有多行
     * @param keys                        对应数据源对象的属性
     * @param datas                       数据源
     * @param callBack，对excel进一步处理,可为null
     * @param dataCallback，数据进行处理,可为null
     * @return
     */
    HSSFWorkbook createWorkBook(String sheetName, List<String[]> titles,
                                String[] keys, List<Map<String, Object>> datas,
                                IExcelCallBack callBack, IDataCallBack dataCallback);


}
