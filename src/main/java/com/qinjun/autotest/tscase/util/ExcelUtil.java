/* Copyright(c), Qin Jun, All right serverd. */
package com.qinjun.autotest.tscase.util;


import com.qinjun.autotest.tscase.exception.ExceptionUtil;
import com.qinjun.autotest.tscase.exception.TSCaseException;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DecimalFormat;

public class ExcelUtil {
    private final static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    public static Workbook getWorkBook(String path) throws TSCaseException {
        FileInputStream fileInputStream = null;
        Workbook workbook = null;
        try {
            fileInputStream = new FileInputStream(path);
            workbook = WorkbookFactory.create(fileInputStream);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        } finally {
            try {
                fileInputStream.close();
            } catch (Exception e) {
                logger.warn("Get exception when close the file:" + e);
            }
        }
        return workbook;
    }

    public static Sheet getSheet(String path, String sheetName) throws TSCaseException {
        Sheet sheet = null;
        try {
            Workbook workbook = getWorkBook(path);
            sheet = workbook.getSheet(sheetName);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return sheet;
    }


    public static Sheet getSheet(Workbook workbook, String sheetName) throws TSCaseException {
        Sheet sheet = null;
        try {
            sheet = workbook.getSheet(sheetName);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return sheet;
    }


    public static Row getRow(String path, String sheetName, int cellRowNum) throws TSCaseException {
        Row row = null;
        try {
            Sheet sheet = getSheet(path, sheetName);
            row = sheet.getRow(cellRowNum);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return row;
    }


    public static Row getRow(Sheet sheet, int cellRowNum) throws TSCaseException {
        Row row = null;
        try {
            row = sheet.getRow(cellRowNum);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return row;
    }

    public static Cell getCell(String path, String sheetName, int cellRowNum, int cellColumnNum) throws TSCaseException {
        Cell cell = null;
        try {
            Row row = getRow(path,sheetName,cellRowNum);
            cell = row.getCell(cellColumnNum);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return cell;
    }


    public static Cell getCell(Row row, int cellColumnNum) throws TSCaseException {
        Cell cell = null;
        try {
            cell = row.getCell(cellColumnNum);
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return cell;
    }

    public static String getCellValue(String path, String sheetName, int cellRowNum, int cellColumnNum) throws TSCaseException {
        String result = null;
        try {
            Cell cell = getCell(path, sheetName, cellRowNum, cellColumnNum);
            DecimalFormat decimalFormat = new DecimalFormat("#");
            CellType cellType = cell.getCellTypeEnum();
            switch (cellType) {
                case STRING:
                    result = (cell == null ? "" : cell.getStringCellValue());
                    break;
                case NUMERIC:
                    Double tmpDouble = (cell == null ? 0 : cell.getNumericCellValue());
                    result = String.valueOf(tmpDouble);
                    if (result.matches("\\d+.[0]*]")) {
                        int index = result.indexOf(".");
                        result = result.substring(0, index);
                    }
                    if (result.matches("^((-?\\d+.?\\d*)[Ee]{1}(\\d+))$")) {
                        result = decimalFormat.format(tmpDouble);
                    }
                    break;
                case BOOLEAN:
                    Boolean tmpBoolean = (cell == null ? false : cell.getBooleanCellValue());
                    result = String.valueOf(tmpBoolean);
                    break;
                case FORMULA:
                    result = (cell == null ? "" : cell.getCellFormula());
                    break;
                case ERROR:
                    result = (cell == null ? "" : Byte.toString(cell.getErrorCellValue()));
                    break;
                case BLANK:
                    result = "";
                    break;
            }
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return result;
    }


    public static String getCellValue(Cell cell) throws TSCaseException {
        String result = null;
        try {
            DecimalFormat decimalFormat = new DecimalFormat("#");
            CellType cellType = cell.getCellTypeEnum();
            switch (cellType) {
                case STRING:
                    result = (cell == null ? "" : cell.getStringCellValue());
                    break;
                case NUMERIC:
                    Double tmp = (cell == null ? 0 : cell.getNumericCellValue());
                    result = String.valueOf(tmp);
                    if (result.matches("\\d+.[0]*]")) {
                        int index = result.indexOf(".");
                        result = result.substring(0, index);
                    }
                    if (result.matches("^((-?\\d+.?\\d*)[Ee]{1}(\\d+))$")) {
                        result = decimalFormat.format(tmp);
                    }
                    break;
                case BOOLEAN:
                    Boolean tmpBoolean = (cell == null ? false : cell.getBooleanCellValue());
                    result = String.valueOf(tmpBoolean);
                    break;
                case FORMULA:
                    result = (cell == null ? "" : cell.getCellFormula());
                    break;
                case ERROR:
                    result = (cell == null ? "" : Byte.toString(cell.getErrorCellValue()));
                    break;
                case BLANK:
                    result = "";
                    break;
            }
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return result;
    }

    public static int getLastRowNum(String path, String sheetName) throws TSCaseException{
        int lastRowNum = -1;
        try {
            Sheet sheet = getSheet(path, sheetName);
            lastRowNum = sheet.getLastRowNum();

        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return lastRowNum;
    }

    public static int getLastRowNum(Sheet sheet) throws TSCaseException{
        int lastRowNum = -1;
        try {
            lastRowNum = sheet.getLastRowNum();

        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return lastRowNum;
    }


    public static int getLastColumnNum(String path, String sheetName) throws TSCaseException{
        int lastColumnNum = -1;
        try {
            Sheet sheet = getSheet(path, sheetName);
            Row row = sheet.getRow(0);
            lastColumnNum = row.getLastCellNum();

        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return lastColumnNum;
    }

    public static int getLastColumnNum(Sheet sheet) throws TSCaseException{
        int lastColumnNum = -1;
        try {
            Row row = sheet.getRow(0);
            lastColumnNum = row.getLastCellNum();
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return lastColumnNum;
    }


    public static String[][] getSheetContent(String path, String sheetName) throws TSCaseException {
        String[][] content = null;
        try {
            Sheet sheet = getSheet(path, sheetName);
            int lastRowNum =getLastRowNum(sheet);
            int lastColumnNum = getLastColumnNum(sheet);

            content = new String[lastRowNum + 1][lastColumnNum];
            for (int i = 0; i <= lastRowNum; i++) {
                Row row = getRow(sheet,i);
                for (int j = 0; j < lastColumnNum; j++) {
                    Cell cell = getCell(row, j);
                    String value = getCellValue(cell);
                    content[i][j] = value;
                }
            }
        } catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        return content;
    }

    public static void saveWorkBook(Workbook workbook,String path) throws TSCaseException {
        try {
            FileOutputStream fileOutputStream = new FileOutputStream(path);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
        }
        catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
    }

    public static void setCellValue(String path, String sheetName,int rowNum, int columnNum, String value) throws TSCaseException {
        try {
            Workbook workbook = getWorkBook(path);
            Sheet sheet = getSheet(workbook,sheetName);
            Row row = getRow(sheet,rowNum);
            Cell cell = getCell(row,columnNum);
            cell.setCellValue(value);
            saveWorkBook(workbook,path);
        }
        catch (Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:" + exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
    }



}
