package TestEnviroment;

import Portabilidad.Preparation;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class Pre_activation extends Preparation {
    private final Clipboard clipboard;
    private String GrupoUsuario;
    private String Cuentas;
    private String AccountNumber;


    public Pre_activation() throws Exception {
        clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    }


    public void EnterToAccount() throws Exception {
        GetClarify().ShowClarify();
        Thread.sleep(300);
        //Press on Select
        getRobot().mouseMove(217, 35);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Press on cuenta
        Thread.sleep(300);
        getRobot().mouseMove(246, 584);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Click On Account Field
        Thread.sleep(600);
        getRobot().mouseMove(538, 239);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Enter the Account Number
        Thread.sleep(300);
        String AccountNumber = "19448853";
        for (char Letter : AccountNumber.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        //Press on Buscar Button
        getRobot().mouseMove(1446, 427);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Select the Account
        Thread.sleep(300);
        getRobot().mouseMove(427, 518);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Press On Abrir Button
        getRobot().mouseMove(1152, 752);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnVPN() throws Exception {
        Thread.sleep(6000);
        getRobot().mouseMove(1472, 204);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void GetGrupoDeUsuario() throws Exception {
        Thread.sleep(9000);
        //Press on Transferencia Fijo
        getRobot().mouseMove(1055, 199);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on AccountField
        Thread.sleep(500);
        getRobot().mouseMove(812, 293);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Enter the Account Number.
        Thread.sleep(500);
        AccountNumber = "29073468";
        for (char Letter : AccountNumber.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        //press on Buscar Button
        Thread.sleep(200);
        getRobot().mouseMove(918, 286);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Clean Message Error
        Thread.sleep(500);
        getRobot().mouseMove(670, 465);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Select Grupo Usuario Result
        Thread.sleep(500);
        getRobot().mouseMove(931, 365);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_C);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(1000);
    }

    public void CopyTheGU() throws Exception {
        Transferable FirstContent = clipboard.getContents(null);
        GrupoUsuario = (String) FirstContent.getTransferData(DataFlavor.stringFlavor);
    }

    public void SearchGU() throws Exception {
        Thread.sleep(300);
        getRobot().mouseMove(811, 202);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1031, 267);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        for (char FirstLetter : GrupoUsuario.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(FirstLetter));
            break;
        }
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void SearchInAnotherGU() throws Exception {
        getRobot().mouseMove(1094, 272);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().keyPress(KeyEvent.VK_DOWN);
        getRobot().keyRelease(KeyEvent.VK_DOWN);
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(600);
    }

    public void CopyCuentasDeGrupoUsuario() throws Exception {
        //Select Cuentas
        Thread.sleep(300);
        getRobot().mouseMove(610, 246);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Copy Cuentas
        Thread.sleep(300);
        getRobot().mousePress(InputEvent.BUTTON3_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
        Thread.sleep(500);
        getRobot().mouseMove(631, 263);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        Transferable GetCuenta = clipboard.getContents(null);
        Cuentas = (String) GetCuenta.getTransferData(DataFlavor.stringFlavor);
        System.out.println(Cuentas);
    }

    public void SearchAndSelectCuenta() throws Exception {
        Pattern CuentaPattern = Pattern.compile("\\d{8}");
        Matcher CuentaMatch = CuentaPattern.matcher(Cuentas);
        int CuentaPosition = -1;
        boolean IsCuenta = false;
        while (CuentaMatch.find()) {
            CuentaPosition++;
            if (CuentaMatch.group().equals(AccountNumber)) {
                getRobot().mouseMove(553, 258);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                for (int i = 1; i <= CuentaPosition; i++) {
                    getRobot().keyPress(KeyEvent.VK_DOWN);
                    getRobot().keyRelease(KeyEvent.VK_DOWN);
                    IsCuenta = true;
                }
            break;
            }
        }
        if (!IsCuenta) {
            SearchInAnotherGU();
            CopyCuentasDeGrupoUsuario();
            SearchAndSelectCuenta();
        }
    }
}
