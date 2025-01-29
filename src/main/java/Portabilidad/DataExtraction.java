package Portabilidad;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.util.LinkedList;


public class DataExtraction {

    private final Sheet RevisionSheet;
    private final Sheet PlanesSheet;
    private final LinkedList<LinkedList<String>> AllData = new LinkedList<>();
    private Cell SfidCell = null;
    private static String Sfid;

    public LinkedList<LinkedList<String>> GetAllData() {
        return AllData;
    }


    public DataExtraction() throws Exception {
        File RevisionFile = new File("C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx");
        try (FileInputStream EnterToFile = new FileInputStream(RevisionFile);
             Workbook operationSheet = new XSSFWorkbook(EnterToFile)) {
            RevisionSheet = operationSheet.getSheet("Table1");
            File PlanesFile = new File("C:\\Portabilidad_Auto_End2End\\data\\PlanGsm.xlsx");
            FileInputStream ReadPlanes = new FileInputStream(PlanesFile);
            Workbook PlanesWorkBook = new XSSFWorkbook(ReadPlanes);
            PlanesSheet = PlanesWorkBook.getSheet("Planes");
        }
    }


    public void ExtractNumeroMovil() throws Exception{
        ExtractSfid();
        int RowNum = 0;
        for (Row row : RevisionSheet) {
            OuterLoop:
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1 && cell.getColumnIndex() != SfidCell.getColumnIndex()) {
                    if (cell.toString().contains("portabilidad") || cell.toString().contains("Portabilidad")) {
                        for (Cell TargetCell : row) {
                            if (TargetCell.toString().matches("^7.?\\d{8}|^6.?\\d{8}")) {
                                FileLogging.GetSheet().createRow(RowNum++).createCell(0).setCellValue(TargetCell.toString());
                                break OuterLoop;
                            }
                        }
                    }
                }
            }
        }
    }

    public void ExtractSfid() {
        outerLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (row.getCell(cell.getColumnIndex()).getRowIndex() >= 1) {
                    if (cell.toString().matches("([^YyZz]\\w?\\d+[A-Z]$)")) {
                        SfidCell = row.getCell(cell.getColumnIndex());
                        Sfid = cell.getStringCellValue();
                        break outerLoop;
                    }
                }
            }
        }
    }
    public static String GetSfid(){
        return Sfid;
    }
}