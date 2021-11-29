package ru.eleron.org;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.wp.usermodel.HeaderFooterType;
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
            XSSFSheet sheet = workbook.getSheet("Лист2");
            Iterator<Row> rowIterator = sheet.rowIterator();
            int cellNumber;
            File docxFile = new File("C:\\Code\\tests.docx");
            file.createNewFile();
            XWPFDocument document = new XWPFDocument();
            document.createHeader(HeaderFooterType.DEFAULT);

            while (rowIterator.hasNext()) {
                XSSFRow row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                cellNumber = 1;
                XWPFTable table = createTable(document, 10, 1);
                while (cellIterator.hasNext()) {
                    Cell next = cellIterator.next();
                    switch (cellNumber) {
                        case 1:
                            document.createParagraph().createRun().setText("\n");
                            table.getCTTbl().newCursor().toEndToken();
                            XWPFParagraph xwpfParagraph = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph.createRun().setText("Тип теста: " + readValue(next).trim());
                            break;
                        case 2:

                            XWPFParagraph xwpfParagraph2 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph2.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph2.createRun().setText("Заголовок теста: " + readValue(next).trim());
                            break;
                        case 3:
                            table.getCTTbl().newCursor().toEndToken();
                            XWPFParagraph xwpfParagraph3 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph3.setAlignment(ParagraphAlignment.CENTER);
                            XWPFRun run = xwpfParagraph3.createRun();
                            run.setBold(true);
                            run.setText("Вопрос: " + readValue(next).trim());
                            break;
                        case 4:
                            break;
                        case 5:
                            String value1 = readValue(next);
                            if (value1.equals("'не задано'")) {
                                break;
                            }
                            table.getCTTbl().newCursor().toFirstContentToken();
                            XWPFParagraph xwpfParagraph5 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph5.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph5.createRun().setText("Ответ 1: " + readValue(next).trim());
                            break;
                        case 6:
                            String value2 = readValue(next);
                            if (value2.equals("'не задано'")) {
                                break;
                            }
                            XWPFParagraph xwpfParagraph6 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph6.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph6.createRun().setText("Ответ 2: " + readValue(next).trim());
                            break;
                        case 7:
                            String value3 = readValue(next);
                            if (value3.equals("'не задано'")) {
                                break;
                            }
                            XWPFParagraph xwpfParagraph7 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph7.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph7.createRun().setText("Ответ 3: " + readValue(next).trim());
                            break;
                        case 8:
                            String value4 = readValue(next);
                            if (value4.equals("'не задано'")) {
                                break;
                            }
                            XWPFParagraph xwpfParagraph8 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph8.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph8.createRun().setText("Ответ 4: " + readValue(next).trim());
                            break;
                        case 9:
                            String value5 = readValue(next);
                            if (value5.equals("'не задано'")) {
                                break;
                            }
                            XWPFParagraph xwpfParagraph9 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph9.setAlignment(ParagraphAlignment.CENTER);
                            xwpfParagraph9.createRun().setText("Ответ 5: " + readValue(next).trim());
                            break;
                        case 10:
                            String value6 = readValue(next);
                            if (value6.equals("'не задано'")) {
                                break;
                            }
                            XWPFParagraph xwpfParagraph10 = table.getRow(cellNumber - 1).getCell(0).addParagraph();
                            xwpfParagraph10.setAlignment(ParagraphAlignment.CENTER);
                            table.getRow(cellNumber - 1).getCell(0).addParagraph().createRun().setText("Ответ 6: " + readValue(next));
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
