package Portabilidad;


import java.io.FileOutputStream;


public class FileLogging extends Check {

    public FileLogging() throws Exception {
    }

    public void SaveFile() throws Exception{
        FileOutputStream SaveFile= new FileOutputStream(DataExtraction.getFile());
        DataExtraction.getWorkbook().write(SaveFile);
    }
    public void CloseFile()throws Exception{
        DataExtraction.getWorkbook().close();
    }
    public void ExecuteFileLoggingMethods() throws Exception{
        SaveFile();
        CloseFile();
    }
}