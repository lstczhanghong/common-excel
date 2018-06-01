package com.leadingsoft.common;

import org.apache.poi.hssf.usermodel.HSSFSheet;

/**
 * excel 表处理
 */
public interface IExcelCallBack {
    void run(HSSFSheet sheet);
}
