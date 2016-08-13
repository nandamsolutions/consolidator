package com.nandamsolutions.consolidator;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookUtils {
    public static Workbook workbook(InputStream stream, boolean isXlsx) throws IOException {
        return isXlsx ? new XSSFWorkbook(stream) : new HSSFWorkbook(stream);
    }
    
    public enum WorkbookType {
        XLS, XLSX;
        
        public static WorkbookType get(String file) {
            if(file.toLowerCase().endsWith(".xlsx")) {
                return WorkbookType.XLSX;
            } else if(file.toLowerCase().endsWith(".xls")) {
                return WorkbookType.XLS;
            }
            return null;
        }
    }

    public static boolean isStringCell(Cell cell) {
        return cell.getCellType() == Cell.CELL_TYPE_STRING;
    }

    public static double getNumericValue(Cell cell) {
        if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
            return Double.parseDouble(cell.getStringCellValue());
        } else if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
            return cell.getNumericCellValue();
        } else if (Cell.CELL_TYPE_BLANK == cell.getCellType()) {
            return 0;
        }
        throw new UnsupportedValueTypeException(cell.getCellType());
    }

    public static class UnsupportedValueTypeException extends RuntimeException {
        private static final long serialVersionUID = 5033417292849095402L;

        public UnsupportedValueTypeException(int cellType) {
            super(String.valueOf(cellType));
        }
    }
}