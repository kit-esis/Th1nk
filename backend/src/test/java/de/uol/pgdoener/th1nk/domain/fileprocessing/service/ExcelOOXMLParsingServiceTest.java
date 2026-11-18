package de.uol.pgdoener.th1nk.domain.fileprocessing.service;

import de.uol.pgdoener.th1nk.domain.fileprocessing.helper.DateNormalizerService;
import de.uol.pgdoener.th1nk.domain.fileprocessing.helper.NumberNormalizerService;
import de.uol.pgdoener.th1nk.domain.fileprocessing.helper.TypeDetector;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class ExcelOOXMLParsingServiceTest {

    private final ExcelOOXMLParsingService service = new ExcelOOXMLParsingService(
            new DateNormalizerService(), new NumberNormalizerService(), new TypeDetector());

    @Test
    void shortRowsArePaddedWithEmptyStrings() throws IOException {
        byte[] xlsx;
        try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet();
            Row full = sheet.createRow(0);
            for (int i = 0; i < 5; i++) full.createCell(i).setCellValue("h" + i);
            Row shortRow = sheet.createRow(1);
            shortRow.createCell(0).setCellValue("a");
            shortRow.createCell(1).setCellValue("b");
            workbook.write(out);
            xlsx = out.toByteArray();
        }

        String[][] result = service.readExcel(new ByteArrayInputStream(xlsx), Optional.empty());

        assertEquals(2, result.length);
        for (String[] row : result) {
            assertEquals(5, row.length);
            for (String cell : row) assertNotNull(cell);
        }
        assertEquals("a", result[1][0]);
        assertEquals("", result[1][4]);
    }
}
