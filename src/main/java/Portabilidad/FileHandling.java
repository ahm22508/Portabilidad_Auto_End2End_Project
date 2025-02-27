package Portabilidad;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class FileHandling {
    private final robot Bot = robot.getBotInstance();

    public FileHandling() throws Exception {

    }

    public void RenameFile() throws Exception {
        String FileName = "Operation";
        Thread.sleep(4000);
        Bot.getRobot().keyPress(KeyEvent.VK_F2);
        Bot.getRobot().keyRelease(KeyEvent.VK_F2);
        Thread.sleep(500);
        for (char letter : FileName.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(letter));
        }
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(100);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_X);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_X);
    }

    public void MoveFile() throws Exception {
        String[] Cd = {"cmd", "/c", "cd C:\\Portabilidad_Auto_End2End\\data && start ."};
        ProcessBuilder proc = new ProcessBuilder();
        proc.command(Cd);
        proc.start();
        Thread.sleep(6000);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_V);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(1000);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void OpenFile() throws Exception {
        Thread.sleep(1000);
        String[] Command = {"cmd", "/c", "C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx"};
        ProcessBuilder proc = new ProcessBuilder();
        proc.command(Command);
        proc.start();
        Thread.sleep(12000);
        Bot.getRobot().mouseMove(794, 90);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
    }

    //Revision Code.
    public String FileIterator() throws IOException {
        String SheetName ="";
        File OperationFile = new File("C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx");
        FileInputStream AccessToFile = new FileInputStream(OperationFile);
        Workbook iterateSheets = new XSSFWorkbook(AccessToFile);
        for(int i = 0; i < iterateSheets.getNumberOfSheets(); i++){
            if(iterateSheets.getSheetName(i).contains("Revision") || iterateSheets.getSheetName(i).contains("sheet1") || iterateSheets.getSheetName(i).contains("revision") || iterateSheets.getSheetName(i).contains("Sheet1")|| iterateSheets.getSheetName(i).contains("REVISION") || iterateSheets.getSheetName(i).contains("SHEET1")){
                SheetName = iterateSheets.getSheetName(i);
            }
        }
        return SheetName;
    }
   // Review Code.
    public void SearchToSheet() throws Exception {
        Thread.sleep(700);
        Bot.getRobot().keyPress(KeyEvent.VK_F5);
        Bot.getRobot().keyRelease(KeyEvent.VK_F5);
        Thread.sleep(500);
        for (char letter : FileIterator().toCharArray()) {
            Thread.sleep(100);
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(letter));
        }
        Bot.getRobot().keyPress(KeyEvent.VK_SHIFT);
        Bot.getRobot().keyPress(KeyEvent.VK_1);
        Bot.getRobot().keyRelease(KeyEvent.VK_SHIFT);
        Bot.getRobot().keyRelease(KeyEvent.VK_1);
        Bot.getRobot().keyPress(KeyEvent.VK_A);
        Bot.getRobot().keyPress(KeyEvent.VK_1);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void TransformFileToText() throws Exception {
        Thread.sleep(1000);
        Bot.getRobot().mouseMove(355, 62);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        Bot.getRobot().mouseMove(90, 140);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(17000);
        Bot.getRobot().mouseMove(510, 240);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_A);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_A);
        Thread.sleep(500);
        Bot.getRobot().mouseMove(577, 58);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().mouseMove(572, 281);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void UpdateAndSaveFile() throws Exception {
        Thread.sleep(1000);
        Bot.getRobot().mouseMove(19, 75);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(9000);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_S);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_S);
        Thread.sleep(800);
        Bot.getRobot().mouseMove(1510, 22);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

   public void ExecuteFileHandlingMethods() throws Exception{
        RenameFile();
        MoveFile();
        OpenFile();
        SearchToSheet();
        TransformFileToText();
        UpdateAndSaveFile();
        Thread.sleep(1000);
   }


}