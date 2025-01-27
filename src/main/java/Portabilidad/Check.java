package Portabilidad;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;


public class Check extends Preparation{


    private final Screen pcScreen;
    private final Pattern RecognizePopUp = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\PopUp.png");
    private final Pattern activationImage = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\ActivationImage.png");

    public Check() throws Exception {
        pcScreen = new Screen();
    }



    public void generalCheck() throws Exception {

        //Press on Select
        Thread.sleep(500);
        getRobot().mouseMove(226, 37);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on Servicios
        Thread.sleep(500);
        getRobot().mouseMove(246, 654);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on DropMenu
        getRobot().mouseMove(452, 252);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on DropDownMenu
        getRobot().mouseMove(455, 288);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnAC() throws Exception {
        //press on AC
        Thread.sleep(300);
        getRobot().mouseMove(350, 270);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnPA() throws Exception {
        // press on PA
        Thread.sleep(300);
        getRobot().mouseMove(346, 280);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnIdentificadorDeServicio() throws Exception {
        //press on Identificador de Servicio
        getRobot().mouseMove(1183, 254);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
    }

    public void ExecuteCheck() throws Exception {
        //press on Buscar Button
        getRobot().mouseMove(1476, 464);
        getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1047);
    }

    public void CheckAc() throws Exception {
        for (Row Lineas : FileLogging.getSheet()) {
            Cell LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                String Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("No Activa");
                //Clean Error.
                getRobot().mouseMove(764, 475);
                getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            if (pcScreen.has(activationImage) && !pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("Activa");
            }
            CleanCheck();
        }
    }

    public void CheckPA() throws Exception {
        for (Row Lineas : FileLogging.getSheet()) {
            Cell LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                String Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("No Preactivada");
                //Clean Error.
                getRobot().mouseMove(764, 475);
                getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            if (!pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(1).setCellValue("Preactivada");
            }
            CleanCheck();
        }
    }
    public void CheckIfIsNotPA() throws Exception{
        for (Row Lineas : FileLogging.getSheet()) {
            Cell LineaCell = Lineas.getCell(0);
            if (LineaCell != null) {
                String Linea = LineaCell.getStringCellValue();
                for (char digit : Linea.toCharArray()) {
                    getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
                }
            }
            ExecuteCheck();
            if (pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(2).setCellValue("No Preactivada(Not_Error)");
                //Clean Error.
                getRobot().mouseMove(764, 475);
                getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            }
            if (!pcScreen.has(RecognizePopUp)) {
                Lineas.createCell(2).setCellValue("Preactivada(Error)");
            }
            CleanCheck();
        }
    }

    public void CleanCheck() throws Exception {
        //Cleaning Field.
        getRobot().mouseMove(1050, 254);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().keyPress(KeyEvent.VK_DELETE);
        getRobot().keyRelease(KeyEvent.VK_DELETE);
        Thread.sleep(200);
    }

    public void EndCheck() {
        //Press on Cerrar Button.
        getRobot().mouseMove(977, 764);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
}