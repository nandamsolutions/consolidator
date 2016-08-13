package com.nandamsolutions.consolidator;

import static com.nandamsolutions.consolidator.WorkbookUtils.getNumericValue;
import static com.nandamsolutions.consolidator.WorkbookUtils.isStringCell;

import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class WorkbookService {
    
    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookService.class);
    
    public Map<String, Double> read(Workbook workbook) {
        Map<String, Double> fieldValues = new LinkedHashMap<>();
        workbook.iterator().forEachRemaining(sheet -> sheet.forEach(row -> {
            Cell labelCell = row.getCell(0);
            if (isStringCell(labelCell)) {
                try {
                    fieldValues.put(labelCell.getStringCellValue(), getNumericValue(row.getCell(1)));
                } catch (Exception e) {
                    LOGGER.error("Failed to read row", e);
                }
            }
        }));
        return fieldValues;
    }
    
    public void write(Map<String, Double> data, OutputStream os) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Total");
            int rowNum = 0;
            for (Entry<String, Double> entry : data.entrySet()) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0, Cell.CELL_TYPE_STRING).setCellValue(entry.getKey());
                row.createCell(1, Cell.CELL_TYPE_NUMERIC).setCellValue(entry.getValue());
            }
            workbook.write(os);
        }
    }
}
