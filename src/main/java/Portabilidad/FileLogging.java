package Portabilidad;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class FileLogging extends Check {
    private static Sheet ResultSheet;
    private final File OpenFile;
    private final Workbook OpenSheet;


    public FileLogging() throws Exception {
        OpenFile = new File("C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx");
        FileInputStream handleFile = new FileInputStream(OpenFile);
        OpenSheet = new XSSFWorkbook(handleFile);
        if(OpenSheet.getSheet("Result") == null) {
            ResultSheet = OpenSheet.createSheet("Result");
        }
        else {
            ResultSheet = OpenSheet.getSheet("Result");
        }
    }
    public static Sheet GetSheet(){
        return ResultSheet;
    }
    public void SaveFile() throws Exception{
        FileOutputStream SaveFile= new FileOutputStream(OpenFile);
        OpenSheet.write(SaveFile);
    }
    public void CloseFile()throws Exception{
        OpenSheet.close();
    }
    public void ExecuteFileLoggingMethods() throws Exception{
        SaveFile();
        CloseFile();
    }
}