package com.example.testapachepoi.service;
import com.example.testapachepoi.entity.Book;
import com.example.testapachepoi.repository.BookRepository;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.Iterator;
import java.util.List;

@Service
public class ImportExcelService {
    private final int COLUMN_INDEX_ID = 0;
    private final int COLUMN_INDEX_TITLE = 1;
    private final int COLUMN_INDEX_PRICE = 2;
    private final int COLUMN_INDEX_QUANTITY = 3;
    private final int COLUMN_INDEX_TOTAL = 4;

    @Autowired
    private BookRepository bookRepository;

    public boolean readExcel(String excelFilePath) throws IOException {
        try {
            // Get file
            InputStream inputStream = new FileInputStream(new File(excelFilePath));

            // Get workbook
            Workbook workbook = getWorkbook(inputStream, excelFilePath);

            // Get sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Get all rows
            Iterator<Row> iterator = sheet.iterator();
            while (((Iterator<?>) iterator).hasNext()) {
                Row nextRow = iterator.next();
                if (nextRow.getRowNum() == 0) {
                    // Ignore header
                    continue;
                }

                // Get all cells
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                // Read cells and set value for book object
                Book book = new Book();
                while (cellIterator.hasNext()) {
                    //Read cell
                    Cell cell = cellIterator.next();
                    Object cellValue = getCellValue(cell);
                    if (cellValue == null || cellValue.toString().isEmpty()) {
                        continue;
                    }
                    // Set value for book object
                    int columnIndex = cell.getColumnIndex();
                    switch (columnIndex) {
                        case COLUMN_INDEX_ID:
                            book.setId(new BigDecimal((double) cellValue).intValue());
                            break;
                        case COLUMN_INDEX_TITLE:
                            book.setTitle((String) getCellValue(cell));
                            break;
                        case COLUMN_INDEX_QUANTITY:
                            book.setQuantity(new BigDecimal((double) cellValue).intValue());
                            break;
                        case COLUMN_INDEX_PRICE:
                            book.setPrice((Double) getCellValue(cell));
                            break;
                        case COLUMN_INDEX_TOTAL:
                            book.setTotalMoney((Double) getCellValue(cell));
                            break;
                        default:
                            break;
                    }

                }
                bookRepository.save(book);
            }

            workbook.close();
            inputStream.close();
            return true;
        }catch (Exception e) {
            System.out.println(e);
            return false;
        }
    }

    // Get Workbook
    private Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Get cell value
    private Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }

        return cellValue;
    }

    public void generateProductStatisticalExcel(HttpServletResponse response) {
        try{
            List<Book> books  = bookRepository.findAll();
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Customer Info");
            HSSFRow row = sheet.createRow(0);

            row.createCell(0).setCellValue("Id");
            row.createCell(1).setCellValue("Tiêu đề");
            row.createCell(2).setCellValue("Giá");
            row.createCell(3).setCellValue("Số lượng");
            row.createCell(4).setCellValue("Tổng tiền");

            int dataRowIndex = 1;

            for (Book book : books) {
                HSSFRow dataRow = sheet.createRow(dataRowIndex);

                dataRow.createCell(0).setCellValue(book.getId());
                dataRow.createCell(1).setCellValue(book.getTitle());
                dataRow.createCell(2).setCellValue(book.getPrice());
                dataRow.createCell(3).setCellValue(book.getQuantity());
                dataRow.createCell(4).setCellValue(book.getTotalMoney());

                dataRowIndex++;
            }
            ServletOutputStream ops = response.getOutputStream();
            workbook.write(ops);
            workbook.close();
            ops.close();
        }
        catch (Exception e){
            System.out.println(e);
        }
    }
}
