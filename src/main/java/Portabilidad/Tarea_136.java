package Portabilidad;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.sikuli.script.FindFailed;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class Tarea_136 {

    private final File Proceso136 = new File("C:\\Portabilidad_Auto_End2End\\data\\Proceso_136.xlsx");
    private final FileInputStream openTheFile = new FileInputStream(Proceso136);
    private final Workbook ProcesoWorkBook = new XSSFWorkbook(openTheFile);
    private final ProcessBuilder Process = new ProcessBuilder();
    private final Screen pcScreen = new Screen();
    private final Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    private final robot Bot = robot.getBotInstance();

    public Tarea_136() throws Exception{
    }


    public String getCif() throws Exception{
        Thread.sleep(1000);
        Bot.getRobot().mouseMove(217, 35);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(284, 232);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().delay(1000);
        Bot.getRobot().mouseMove(466, 316);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for(char Letter : DataExtraction.GetCuenta().toCharArray()){
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        Bot.getRobot().mouseMove(1466, 485);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().delay(1000);
        Bot.getRobot().mouseMove(452 , 544);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Transferable Content = clipboard.getContents(null);
        Bot.getRobot().mouseMove(1481 , 760);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

        return (String)Content.getTransferData(DataFlavor.stringFlavor);
    }

    public Sheet Hoja1Sheet(){

        return ProcesoWorkBook.getSheet("Hoja1");
    }
    public void PopulateSheet(){
        int i = 0;
        for (Row row : DataExtraction.GetResultSheet()) {
            Cell ResultCell = row.getCell(2);
            if (ResultCell.getStringCellValue().equals("No")) {
                i++;
                Hoja1Sheet().createRow(i).createCell(0).setCellValue("B29733870");
                Hoja1Sheet().getRow(i).createCell(1).setCellValue("600504056");
                Hoja1Sheet().getRow(i).createCell(2).setCellValue("24759476");
                Hoja1Sheet().getRow(i).createCell(3).setCellValue("29584847");
                Hoja1Sheet().getRow(i).createCell(4).setCellValue("6001300351760");
                Hoja1Sheet().getRow(i).createCell(5).setCellValue("Integrado");
                Hoja1Sheet().getRow(i).createCell(6).setCellValue("MPMVB");
                Hoja1Sheet().getRow(i).createCell(7).setCellValue("ORo");
                Hoja1Sheet().getRow(i).createCell(8).setCellValue("150");
                Hoja1Sheet().getRow(i).createCell(12).setCellValue("Error Preactivacion");
                Hoja1Sheet().getRow(i).createCell(13).setCellValue("Error Preactivacion");
                Hoja1Sheet().getRow(i).createCell(14).setCellValue("Error");
                Hoja1Sheet().getRow(i).createCell(15).setCellValue("A4383945S");
            }
        }
    }
    public void ChangeEXT() throws Exception{
        String [] Commands = {"cmd" , "/c" ,"C:\\Portabilidad_Auto_End2End\\data\\Script.vbs"};
        Process.command(Commands);
        Process.start();
    }
    public void SaveFile() throws Exception{
        FileOutputStream FileSaving  = new FileOutputStream(Proceso136);
        ProcesoWorkBook.write(FileSaving);
    }
    public void closeFile()throws Exception{
        ProcesoWorkBook.close();
    }
    public void OpenCSVfile() throws Exception{
        Thread.sleep(500);
        String [] OpenCommand = {"cmd","/c","C:\\Portabilidad_Auto_End2End\\data\\Hoja1.csv"};
        Process.command(OpenCommand);
        Process.start();
    }
    public void ChangeToText() throws Exception{
        String Text = "Text";
        Thread.sleep(8000);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(2000);
        Bot.getRobot().mouseMove(749 , 476);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_A);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_A);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(794 , 100);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for(char Letter : Text.toCharArray()){
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(500);
        Bot.getRobot().mouseMove(323 , 262);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void CopyFileContent()throws Exception{
        Thread.sleep(200);
        Bot.getRobot().mouseMove(52 , 245);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_A);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_A);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Bot.getRobot().keyPress(KeyEvent.VK_ALT);
        Bot.getRobot().keyPress(KeyEvent.VK_F4);
        Bot.getRobot().keyRelease(KeyEvent.VK_ALT);
        Bot.getRobot().keyRelease(KeyEvent.VK_F4);
        Thread.sleep(1000);
        Bot.getRobot().keyPress(KeyEvent.VK_TAB);
        Bot.getRobot().keyRelease(KeyEvent.VK_TAB);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }
    public void OpenWindowsExplorer() throws FindFailed {
        Pattern WindowExplorer = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\Windows Explorer.png");
        pcScreen.click(WindowExplorer);
    }
    public void OpenSheet() throws Exception{
        Thread.sleep(1000);
         Pattern SquareImg = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\Square.png");
         pcScreen.click(SquareImg);
         Thread.sleep(1000);
         Pattern AfouadsDir = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\AfouadsDir.png");
         pcScreen.doubleClick(AfouadsDir);
         Thread.sleep(800);
        Bot.getRobot().mouseMove(374 , 322);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
    }
    public void ToText() throws Exception{
        Thread.sleep(4500);
        Bot.getRobot().mouseMove(21 , 197);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(777 , 69);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        Bot.getRobot().mouseMove(710 , 548);
        Thread.sleep(800);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
    public void PasteContent() throws Exception{
        Thread.sleep(200);
        Bot.getRobot().mouseMove(50 , 208);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_V);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_V);
    }
    public void saveSheet() throws Exception{
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_G);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_G);
        Thread.sleep(400);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_ALT);
        Bot.getRobot().keyPress(KeyEvent.VK_F4);
        Bot.getRobot().keyRelease(KeyEvent.VK_ALT);
        Bot.getRobot().keyRelease(KeyEvent.VK_F4);
        Thread.sleep(100);
        Bot.getRobot().keyPress(KeyEvent.VK_TAB);
        Bot.getRobot().keyRelease(KeyEvent.VK_TAB);
        Thread.sleep(100);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }
    public void RestoreCitrixExplorer()throws Exception{
        Pattern PC = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\PC.png");
        pcScreen.click(PC);
        Thread.sleep(400);
        Bot.getRobot().mouseMove(511 , 517);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Pattern Minimize = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\Minimize.png");
        pcScreen.click(Minimize);
    }
    public void uploadTheSheet() throws Exception{
        Windows windows = new Windows();
        windows.ShowWindow("more");
        Thread.sleep(1000);
        Pattern Gescore = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\Gescore.png");
        pcScreen.click(Gescore);
        Thread.sleep(2000);
        Bot.getRobot().mouseMove(600 , 366);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(5000);
        Bot.getRobot().mouseMove(230 , 98);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(2000);
        Bot.getRobot().mouseMove(254 , 178);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        Bot.getRobot().mouseMove(614 , 591);
        for(int i = 0; i < 4; i++){
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        Thread.sleep(500);
        Bot.getRobot().mouseMove(274 , 594);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(728 , 226);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(5000);
        Bot.getRobot().mouseMove(522 , 294);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(654 , 188);
        Pattern Maximize = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\Maximize.png");
        pcScreen.mouseMove(Maximize);
        Thread.sleep(600);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);

        for(int i =0; i < 60; i++){
            Thread.sleep(30);
            Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
            Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
        }
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);


        Pattern RecycleBin = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\RecycleBin.png");
        pcScreen.click(RecycleBin);
        Pattern PC = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\PC.png");
        pcScreen.click(PC);
        Bot.getRobot().mouseMove(336 , 546);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(400);
        Pattern Afouads = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\AfouadsDir.png");
        pcScreen.doubleClick(Afouads);
        Thread.sleep(1500);
        Bot.getRobot().mouseMove(507 , 46);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        String Hoja1 = "Hoja1";
        for(char Letter : Hoja1.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        Thread.sleep(1500);
        Bot.getRobot().mouseMove(258 , 124);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(2500);
        Bot.getRobot().mouseMove(931 , 351);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(6000);
        Bot.getRobot().mouseMove(951 , 304);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(6000);
        Bot.getRobot().mouseMove(532 , 202);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(600);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(900);
    }
    public String getIDCarga() throws Exception{
        String idCarga = "";
        try {
            Transferable FirContent = clipboard.getContents(null);
             idCarga =(String) FirContent.getTransferData(DataFlavor.stringFlavor);
        }
        catch (IllegalStateException i){
            i.getCause();
        }
        return idCarga ;
    }
    public void refreshAndClean() throws Exception{
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_R);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_R);
        Thread.sleep(6000);
        Bot.getRobot().mouseMove(721 , 92);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void executeTarea_135methods()throws Exception{
        PopulateSheet();
        SaveFile();
        closeFile();
        ChangeEXT();
        OpenCSVfile();
        ChangeToText();
        CopyFileContent();
        OpenWindowsExplorer();
        OpenSheet();
        ToText();
        PasteContent();
        saveSheet();
        RestoreCitrixExplorer();
        Thread.sleep(2000);
        uploadTheSheet();
        refreshAndClean();
    }
}