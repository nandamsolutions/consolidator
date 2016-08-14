package com.nandamsolutions.consolidator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookService {

    public void mergeBooks(Collection<Workbook> workbooks, Workbook consolidated) {
        List<List<Row>> workbooksRows = workbooks.stream().map(this::rowsToConsolidate).collect(Collectors.toList());
        List<Row> consolidatedRows = this.rowsToConsolidate(consolidated);
        for (List<Row> rows : workbooksRows) {
            if (consolidatedRows.size() != rows.size()) {
                throw new IllegalStateException(
                        "Number of data rows doesn't match to consolidate. Please fix input workbooks");
            }
        }

        for (int rowNumber = 0; rowNumber < consolidatedRows.size(); rowNumber++) {
            Row row = consolidatedRows.get(rowNumber);
            for (int cellNumber = 0; cellNumber < row.getLastCellNum(); cellNumber++) {
                Cell cell = row.getCell(cellNumber);
                if (cell != null && cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    int rn = rowNumber, cn = cellNumber;
                    cell.setCellValue(workbooksRows.stream().map(rows -> rows.get(rn).getCell(cn).getNumericCellValue())
                            .reduce(0.0, (r, v) -> r + v));
                }
            }
        }
    }

    private List<Row> rowsToConsolidate(Workbook workbook) {
        Iterator<Row> iterator = workbook.getSheetAt(0).iterator();
        List<Row> rows = new ArrayList<>();
        while (iterator.hasNext()) {
            Row row = iterator.next();
            if (row.getCell(0) != null && row.getCell(0).getCellType() == Cell.CELL_TYPE_STRING
                    && row.getCell(0).getStringCellValue().trim().equalsIgnoreCase("MH-2038")) {
                break;
            }
        }
        while (iterator.hasNext()) {
            rows.add(iterator.next());
        }
        return rows;
    }

    public Workbook createCopy(Workbook workbook, String name) throws IOException {
        File tempFile = Files.createTempFile("temp-", ".xls").toFile();
        try (FileOutputStream fos = new FileOutputStream(tempFile)) {
            workbook.write(fos);
        }
        HSSFWorkbook copy = new HSSFWorkbook(new FileInputStream(tempFile));
        FileUtils.deleteQuietly(tempFile);
        copy.getSheetAt(0).getRow(0).getCell(0).setCellValue(name);
        return copy;
    }
}
