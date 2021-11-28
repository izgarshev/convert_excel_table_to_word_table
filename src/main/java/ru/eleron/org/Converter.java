package ru.eleron.org;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

public class Converter {
    private static XSSFWorkbook workbook;

    public static void main(String[] args) {
        File file = new File("C:\\Code\\tests.xlsx");
        if (!file.exists()) {
            System.out.println("файл не существует");
            return;
        }
        System.out.println("файл на месте");
        Set<CellType> types = new HashSet<>();
        try {
            workbook = (XSSFWorkbook) WorkbookFactory.create(file);
            XSSFSheet sheet = workbook.getSheet("Лист1");
            Iterator<Row> rowIterator = sheet.rowIterator();
            int cellNumber;
            try {
//                XWPFParagraph paragraph = document.createParagraph();
//                table.createRow().createCell().addParagraph().createRun().setText("text in the table");
//                paragraph.createRun();
            } catch (Exception e) {
                e.printStackTrace();
            }
            File docxFile = new File("C:\\Code\\tests.docx");
            file.createNewFile();
            XWPFDocument document = new XWPFDocument();
            while (rowIterator.hasNext()) {
                XSSFRow row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                cellNumber = 1;
                XWPFTable table = createTable(document, 7, 1);
                while (cellIterator.hasNext()) {
                    Cell next = cellIterator.next();
                    switch (cellNumber) {
                        case 1:
                            document.createParagraph().createRun().setText("\n");
                            table.getCTTbl().newCursor().toEndToken();
                            XWPFParagraph xwpfParagraph = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph.createRun().setText(readValue(next));
                            break;
                        case 2:
                        case 3:
                        case 4:
                        case 5:
                        case 6:
                        case 7:
                            XWPFParagraph xwpfParagraph2 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph2.setAlignment(ParagraphAlignment.CENTER);
                            table.getRow(cellNumber - 1).getCell(0).addParagraph().createRun().setText(readValue(next));
                            break;
                    }
                    types.add(next.getCellTypeEnum());
                    cellNumber++;
                }
            }
            FileOutputStream fileOutputStream = new FileOutputStream(docxFile);
            document.write(fileOutputStream);
            fileOutputStream.flush();
            fileOutputStream.close();
            System.out.println(types.isEmpty() ? "types is empty" : "types: " + types);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public static String readValue(Cell cell) {
        if (cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
            return String.valueOf(cell.getNumericCellValue());
        } else {
            return cell.getStringCellValue();
        }
    }

    public static XWPFTable createTable(XWPFDocument document, int rows, int cols) {
        return document.createTable(rows, cols);
    }
}
