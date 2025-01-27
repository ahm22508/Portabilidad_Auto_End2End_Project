package NewFunctionality;

import Portabilidad.Check;

import java.awt.event.KeyEvent;

public class FileHandling extends Check {


    public FileHandling() throws Exception {

    }

    public void RenameFile() throws Exception {
        String FileName = "Operation";
        Thread.sleep(4000);
        getRobot().keyPress(KeyEvent.VK_F2);
        getRobot().keyRelease(KeyEvent.VK_F2);
        Thread.sleep(500);
        for (char letter : FileName.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(letter));
        }
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(100);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_X);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_X);
    }

    public void MoveFile() throws Exception {
        String[] Cd = {"cmd", "/c", "cd C:\\Portabilidad_Auto_End2End\\data && start ."};
        ProcessBuilder proc = new ProcessBuilder();
        proc.command(Cd);
        proc.start();
        Thread.sleep(6000);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_V);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(300);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void OpenFile() throws Exception {
        Thread.sleep(1000);
        String[] Command = {"cmd", "/c", "C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx"};
        ProcessBuilder proc = new ProcessBuilder();
        proc.command(Command);
        proc.start();
        Thread.sleep(11000);
        getRobot().mouseMove(794, 90);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
    }

    public void SearchToSheet() throws Exception {
        Thread.sleep(700);
        String SheetName = "Sheet1";
        getRobot().keyPress(KeyEvent.VK_F5);
        getRobot().keyRelease(KeyEvent.VK_F5);
        Thread.sleep(500);
        for (char letter : SheetName.toCharArray()) {
            Thread.sleep(100);
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(letter));
        }
        getRobot().keyPress(KeyEvent.VK_SHIFT);
        getRobot().keyPress(KeyEvent.VK_1);
        getRobot().keyRelease(KeyEvent.VK_SHIFT);
        getRobot().keyRelease(KeyEvent.VK_1);
        getRobot().keyPress(KeyEvent.VK_A);
        getRobot().keyPress(KeyEvent.VK_1);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);


    }

    public void TransformFileToText() throws Exception {
        Thread.sleep(1000);
        getRobot().mouseMove(355, 62);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        getRobot().mouseMove(90, 140);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(17000);
        getRobot().mouseMove(510, 240);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_A);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_A);
        Thread.sleep(500);
        getRobot().mouseMove(577, 58);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().mouseMove(572, 281);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void UpdateAndSaveFile() throws Exception {
        Thread.sleep(1000);
        getRobot().mouseMove(19, 75);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(8000);
        getRobot().mouseMove(1510, 22);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(200);

    }
}