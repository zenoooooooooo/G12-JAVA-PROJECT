package actualsf2practice.Algorithms;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import actualsf2practice.Dashboard;
import actualsf2practice.Interfaces.MethodsInterface;


//All algorithms in this class are developed by E-Jhay Esplana
 public abstract class Methods implements MethodsInterface {

    private static final String path = Dashboard.getSelectedFilePath().toString();
    private static FileInputStream inputStream;
    private static XSSFWorkbook workbook;
    private static XSSFSheet sheet;

    public static void countAbsences(int startRow, int endRow, int startColumn, int endColumn, int absenceCellRow) {
        
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet("Sheet1");

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);

                if (row != null) {
                    Cell cell = row.getCell(1);

                    if (cell != null && !cell.toString().equals("")) {
                        System.out.println(cell.toString());

                        int absences = 0;

                        for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                            Cell eachCell = row.getCell(columnIndex);

                            if (eachCell != null && eachCell.toString().equalsIgnoreCase("x")) {
                                absences++;
                                System.out.println(eachCell.toString());
                            }
                        }

                        System.out.println("Total absences: " + absences);

                        Cell absenceCell = row.getCell(absenceCellRow);
                        if (absenceCell == null) {
                            absenceCell = row.createCell(absenceCellRow);
                        }
                        absenceCell.setCellValue(absences);
                    } else {
                        break;
                    }
                } else {
                    System.out.println("null");
                }
            }

            try (FileOutputStream fileout = new FileOutputStream(path)) {
                workbook.write(fileout);
            }

            workbook.close();
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error reading Excel file: " + ex.getMessage());
        }

    }


    public static void countAbsencesPerDay(int startColumn, int endColumn, int startRow, int endRow, int absenceCellBox) {
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet("Sheet1");

            for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                int totalAbsencesPerDay = 0;

                for (int eachRowIndex = startRow; eachRowIndex <= endRow; eachRowIndex++) {
                    Row eachRow = sheet.getRow(eachRowIndex);
                    Cell eachCell = eachRow.getCell(columnIndex);

                    if (eachCell != null && eachCell.toString().equalsIgnoreCase("x")) {
                        totalAbsencesPerDay++;
                    }
                }

                Row absenceRow = sheet.getRow(absenceCellBox);
                if (absenceRow == null) {
                    absenceRow = sheet.createRow(absenceCellBox);
                }

                Cell absenceCell = absenceRow.getCell(columnIndex);
                if (absenceCell == null) {
                    absenceCell = absenceRow.createCell(columnIndex);
                }

                absenceCell.setCellValue(totalAbsencesPerDay);

            }

            try (FileOutputStream fileout = new FileOutputStream(path)) {
                workbook.write(fileout);
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error reading/writing Excel file: " + ex.getMessage());
        }

    }


    public static void countTotal(int startRow, int startColumn, int endColumn, int absenceCellBox) {
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet("Sheet1");
            int overallTotal = 0;

            for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                Row row = sheet.getRow(startRow);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);

                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        overallTotal += cell.getNumericCellValue();
                    }
                }
            }

            System.out.println(overallTotal);
            Row row = sheet.getRow(startRow);
            Cell cell = row.getCell(absenceCellBox);

            cell.setCellValue(overallTotal);

            try (FileOutputStream fileout = new FileOutputStream(path)) {
                workbook.write(fileout);
            }

        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error reading Excel file: " + ex.getMessage());
        }
    }

    
    public static void countTotalPerDay( int startColumn, int endColumn, int row1, int row2, int row3) {
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet("Sheet1");

            Row r1 = sheet.getRow(row1);
            Row r2 = sheet.getRow(row2);
            Row r3 = sheet.getRow(row3);

            if (r1 != null && r2 != null && r3 != null) {
                for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++) {
                    Cell cell1 = r1.getCell(columnIndex);
                    Cell cell2 = r2.getCell(columnIndex);
                    Cell cell3 = r3.getCell(columnIndex);

                    int value1 = (cell1 != null && cell1.getCellType() == CellType.NUMERIC) ? (int) cell1.getNumericCellValue() : 0;
                    int value2 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? (int) cell2.getNumericCellValue() : 0;

                    int sum = value1 + value2;

                    if (cell3 == null) {
                        cell3 = r3.createCell(columnIndex, CellType.NUMERIC);
                    }
                    cell3.setCellValue(sum);

                }
            }

            try (FileOutputStream fileout = new FileOutputStream(path)) {
                workbook.write(fileout);
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error reading/writing Excel file: " + ex.getMessage());
        }

    }

    public static void countOverallTotal(int row1, int row2, int row3, int column) {
        try {
            inputStream = new FileInputStream(path);
            workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheet("Sheet1");

            Row r1 = sheet.getRow(row1);
            Row r2 = sheet.getRow(row2);
            Row r3 = sheet.getRow(row3);

            Cell cell1 = r1.getCell(column);
            Cell cell2 = r2.getCell(column);
            Cell cell3 = r3.getCell(column);

            int value33 = (cell1 != null && cell1.getCellType() == CellType.NUMERIC) ? (int) cell1.getNumericCellValue() : 0;
            int value59 = (cell2 != null && cell2.getCellType() == CellType.NUMERIC) ? (int) cell2.getNumericCellValue() : 0;

            int sum = value33 + value59;

            cell3.setCellValue(sum);
            try (FileOutputStream fileout = new FileOutputStream(path)) {
                workbook.write(fileout);
            }
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(null, "Error reading/writing Excel file: " + ex.getMessage());
        }
    }

    }

