package Portabilidad;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;


public class Check{


    private final Screen pcScreen = new Screen();
    private final Pattern RecognizePopUp = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\PopUp.png");
    private final Pattern activationImage = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\ActivationImage.png");
    private final robot Bot = robot.getBotInstance();

    public Check() throws Exception {
    }


    public void generalCheck() throws Exception {

        //Press on Select
        Thread.sleep(500);
        Bot.getRobot().mouseMove(226, 37);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on Servicios
        Thread.sleep(500);
        Bot.getRobot().mouseMove(246, 654);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on DropMenu
        Bot.getRobot().mouseMove(452, 252);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on DropDownMenu
        Bot.getRobot().mouseMove(455, 288);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnAC() throws Exception {
        //press on AC
        Thread.sleep(300);
        Bot.getRobot().mouseMove(350, 270);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnPA() throws Exception {
        // press on PA
        Thread.sleep(300);
        Bot.getRobot().mouseMove(346, 280);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnIdentificadorDeServicio() throws Exception {
        //press on Identificador de Servicio
        Bot.getRobot().mouseMove(1183, 254);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
    }

    public void ExecuteCheck() throws Exception {
        //press on Buscar Button
        Bot.getRobot().mouseMove(1476, 464);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1090);
    }

    public void CheckAc() throws Exception {
        Cell LineaCell;
        String Linea;
        for (Row Lineas : DataExtraction.GetResultSheet()) {
             LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                 Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(activationImage) && !pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("Activa");
            }
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("No Activa");
                //Clean Error.
                Bot.getRobot().mouseMove(764, 475);
                Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            CleanCheck();
        }
    }

    public void CheckPA() throws Exception {
        Cell LineaCell;
        String Linea;
        for (Row Lineas : DataExtraction.GetResultSheet()) {
             LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                 Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("No Preactivada");
                //Clean Error.
                Bot.getRobot().mouseMove(764, 475);
                Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            if (!pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("Preactivada");
            }
            CleanCheck();
        }
    }
    public void CheckIfIsNotPA() throws Exception{
        Cell LineaCell;
        String Linea;
        for (Row Lineas : DataExtraction.GetResultSheet()) {
             LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                 Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(2).setCellValue("No Preactivada(Not_Error)");
                //Clean Error.
                Bot.getRobot().mouseMove(764, 475);
                Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            if (!pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(2).setCellValue("Preactivada(Error)");
            }
            CleanCheck();
        }
    }

    public void CleanCheck() throws Exception {
        //Cleaning Field.
        Bot.getRobot().mouseMove(1050, 254);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().keyPress(KeyEvent.VK_DELETE);
        Bot.getRobot().keyRelease(KeyEvent.VK_DELETE);
        Thread.sleep(200);
    }

    public void EndCheck() throws Exception{
        //Press on Cerrar Button.
        Bot.getRobot().mouseMove(977, 764);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void ExecuteCheckPAMethods() throws Exception{
        generalCheck();
        PressOnPA();
        PressOnIdentificadorDeServicio();
        CheckPA();
        EndCheck();
    }
    public void ExecuteCheckACMethods() throws Exception{
         generalCheck();
         PressOnPA();
         PressOnIdentificadorDeServicio();
         CheckIfIsNotPA();
         EndCheck();
         generalCheck();
         PressOnAC();
         PressOnIdentificadorDeServicio();
         CheckAc();
         EndCheck();
    }
}