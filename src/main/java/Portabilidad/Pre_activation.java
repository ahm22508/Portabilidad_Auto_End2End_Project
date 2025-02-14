package Portabilidad;

import dev.failsafe.function.CheckedRunnable;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.sikuli.script.Screen;
import org.sikuli.script.Pattern;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.util.regex.*;



public class Pre_activation extends Preparation {
    private final Clipboard clipboard;
    private String GrupoUsuario;
    private String Cuentas;
    private String AccountNumber;
    private static String LineaError;
    private final Screen SysScreen = new Screen();
    private final Pattern ErrorReservaImage = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\ErrorReservaLinea.png");

    public Pre_activation() throws Exception {
        clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    }

    public static String getLineaError() {
        return LineaError;
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
        for (char Letter : AccountNumber.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        //press on Buscar Button
        Thread.sleep(200);
        getRobot().mouseMove(918, 286);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Clean Message Error
        Thread.sleep(2000);
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
    }

    @SuppressWarnings({"probably busy-waiting", "BusyWait"})
    public void SearchAndSelectCuenta() throws Exception {
        java.util.regex.Pattern CuentaPattern = java.util.regex.Pattern.compile("\\d{8}");
        Matcher CuentaMatch = CuentaPattern.matcher(Cuentas);
        int CuentaPosition = -1;
        boolean IsCuenta = false;
        while (CuentaMatch.find()) {
            CuentaPosition++;
            if (CuentaMatch.group().equals(AccountNumber)) {
                Thread.sleep(50);
                getRobot().mouseMove(553, 258);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(50);
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

    public void SelectPortabilidad() throws Exception {
        getRobot().mouseMove(724, 389);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        getRobot().mouseMove(653, 461);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void IntroduceTheLineNumber() throws Exception {
        getRobot().mouseMove(904, 393);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        for (char Letter : LineaError.toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
    }

    public void IntroduceExtension() throws Exception {
        getRobot().mouseMove(1069, 625);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        for (char Letter : DataExtraction.ExtractExtension().toCharArray()) {
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
    }

    public void PressOnContinuar() {
        getRobot().mouseMove(971, 663);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
    public void FixErrorReserva() throws Exception{
        getRobot().mouseMove(669 , 466);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(556 , 200);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(600);
        getRobot().mouseMove(726 , 320);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_C);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(200);
        getRobot().mouseMove(778 , 743);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_V);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(200);
        getRobot().mouseMove(886 , 735);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(15000);
        getRobot().mouseMove(62 , 137);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        getRobot().mouseMove(113 , 279);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(1286 , 487);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(2000);
        getRobot().mouseMove(355 , 274);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for(char Letter : LineaError.toCharArray()){
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_C);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(200);
        getRobot().mouseMove(563 , 273);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_V);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(200);
        getRobot().mouseMove(582 , 324);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        String LinePrefix = "";
        String LineRango = "";
        for(int i = 0; i < 5; i++){
            if (i <= 2) {
                LinePrefix += LineaError.charAt(i);
            }
            if(i >= 3){
                LineRango += LineaError.charAt(i);
            }
        }
        System.out.println(LineRango);
        System.out.println(LinePrefix);
        if(LinePrefix.startsWith("5")){
            getRobot().keyPress(KeyEvent.VK_DOWN);
            getRobot().keyRelease(KeyEvent.VK_DOWN);
        }
        if(LinePrefix.startsWith("6")) {
            int Prefix = Integer.parseInt(LinePrefix);
            if (Prefix > 605) {
                getRobot().mouseMove(585 , 435);
                for (int coun = 0; coun < Prefix - 605; coun++) {
                    Thread.sleep(50);
                    getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                }
                getRobot().mouseMove(535 , 439);
                Thread.sleep(100);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }
        }
        if(LinePrefix.startsWith("600")){
            getRobot().mouseMove(531 , 370);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("601")){
            getRobot().mouseMove(531 , 382);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("602")) {
            getRobot().mouseMove(531, 399);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("603")){
            getRobot().mouseMove(531 , 423);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("604")){
            getRobot().mouseMove(531 , 419);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("605")){
            getRobot().mouseMove(531 , 435);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }

        if(LinePrefix.startsWith("711")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("717")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("722")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("727")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("744")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
        }

        if(LinePrefix.startsWith("747")){
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
            getRobot().keyRelease(KeyEvent.VK_7);
            getRobot().keyPress(KeyEvent.VK_7);
        }
        getRobot().mouseMove(817 , 324);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        if(LineRango.equals("56")){
            getRobot().keyPress(KeyEvent.VK_DOWN);
            getRobot().keyRelease(KeyEvent.VK_DOWN);
        }
        if(LineRango.equals("00")){
            getRobot().mouseMove(784 , 345);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("01")){
            getRobot().mouseMove(784 , 356);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("02")){
            getRobot().mouseMove(784 , 370);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("03")){
            getRobot().mouseMove(784 , 384);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("04")){
            getRobot().mouseMove(784 , 395);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("05")){
            getRobot().mouseMove(784 , 410);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("06")){
            getRobot().mouseMove(784 , 422);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("07")){
            getRobot().mouseMove(784 , 435);
            Thread.sleep(100);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        int Range = Integer.parseInt(LineRango);
        if(Range > 7){
            getRobot().mouseMove(817, 437);
            for(int i =0; i < Range-7; i++){
                Thread.sleep(50);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }
            getRobot().mouseMove(793, 437);
            getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        Thread.sleep(200);
        getRobot().mouseMove(993, 324);
//        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
//        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        getRobot().mouseMove(335, 423);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1475, 597);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(1481, 760);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(1269, 760);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(65, 135);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        getRobot().mouseMove(101, 264);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(600);
        getRobot().mouseMove(820, 206);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        PressOnContinuar();
    }

    public void IntroduceSfid() throws Exception {
        Thread.sleep(7000);
        String Sfid = DataExtraction.GetSfid();
        int i = 1;
        getRobot().mouseMove(320, 197);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for (char Letter : Sfid.toCharArray()) {
            i++;
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
            if (Sfid.length() - i == 1) {
                getRobot().mouseMove(426, 197);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(100);
                getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Sfid.toCharArray()[i]));
            }
        }
    }

    public void CleanScreen() throws Exception {
        getRobot().mouseMove(1449, 756);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1452, 767);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1272, 758);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1454, 752);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void IntroduceTarjetaSim() throws Exception {
        int i = 0;
        getRobot().mouseMove(333, 261);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for (char Letter : DataExtraction.extractTarjetaSIM().toCharArray()) {
            i++;
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
            if (DataExtraction.extractTarjetaSIM().length() - i == 1) {
                getRobot().mouseMove(426, 260);
                getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(100);
                getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(DataExtraction.extractTarjetaSIM().toCharArray()[i]));
            }
        }
    }

    public void SelectRoaming() throws Exception {
        Thread.sleep(100);
        getRobot().mouseMove(1467, 382);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mouseMove(1320, 438);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void SelectCatalogo() throws Exception {
        String Catalogo = DataExtraction.extractCatalagoDePlan();

        getRobot().mouseMove(1116, 444);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        if (Catalogo.contains("Integrados")) {
            for (int i = 0; i < 6; i++) {
                getRobot().keyPress(KeyEvent.VK_DOWN);
                getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        if (Catalogo.contains("Planes Especial Sol. Empresa VPN")) {
            for (int i = 0; i < 2; i++) {
                getRobot().keyPress(KeyEvent.VK_DOWN);
                getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        if (Catalogo.contains("Planes Soluciones Empresa VPN")) {
            for (int i = 0; i < 4; i++) {
                getRobot().keyPress(KeyEvent.VK_DOWN);
                getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        getRobot().keyPress(KeyEvent.VK_ENTER);
        getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(200);
        getRobot().mouseMove(1210, 466);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    @SuppressWarnings("BusyWait")
    public void SelectPlan() throws Exception {
        Thread.sleep(4000);
        getRobot().mouseMove(869, 490);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().mousePress(InputEvent.BUTTON3_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
        Thread.sleep(900);
        getRobot().mouseMove(900, 568);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        Transferable Planes = clipboard.getContents(null);
        String PlanesEnCatalogo = (String) Planes.getTransferData(DataFlavor.stringFlavor);


        int CoordinateX = 861;
        int CoordinateY = 494;

        java.util.regex.Pattern PlanPattern = java.util.regex.Pattern.compile("[A-Z]{5}");
        Matcher PlanMatch = PlanPattern.matcher(PlanesEnCatalogo);
        getRobot().mouseMove(861, 494);
        Thread.sleep(1000);
        while (PlanMatch.find()) {
            if (CoordinateY < 615) {
                CoordinateY += 14;
                if (PlanMatch.group().equals(DataExtraction.getPlanDeGSM())) {

                    getRobot().mouseMove(CoordinateX, CoordinateY);
                    getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                    Thread.sleep(500);
                    getRobot().keyPress(KeyEvent.VK_ENTER);
                    getRobot().keyRelease(KeyEvent.VK_ENTER);
                    break;
                }
            }

            if (CoordinateY > 615) {
                if (PlanMatch.group().equals(DataExtraction.getPlanDeGSM())) {
                    Thread.sleep(100);
                    getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                    getRobot().keyPress(KeyEvent.VK_ENTER);
                    getRobot().keyRelease(KeyEvent.VK_ENTER);
                    break;
                } else {
                    Thread.sleep(100);
                    getRobot().mouseMove(CoordinateX, CoordinateY);
                    getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                }
            }
        }
    }

    public void SelectPerfilDeConsumo()throws Exception{
        Thread.sleep(2500);
        getRobot().mouseMove(306 , 711);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        if(DataExtraction.extractPerfilDeConsumo().contains("10")){
            getRobot().keyPress(KeyEvent.VK_1);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("1")){
            getRobot().keyPress(KeyEvent.VK_1);
            getRobot().keyPress(KeyEvent.VK_1);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("20")){
            getRobot().keyPress(KeyEvent.VK_2);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("2") || DataExtraction.extractPerfilDeConsumo().contains("plata")){
            getRobot().keyPress(KeyEvent.VK_2);
            getRobot().keyPress(KeyEvent.VK_2);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("300")){
            getRobot().keyPress(KeyEvent.VK_3);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("3")){
            getRobot().keyPress(KeyEvent.VK_3);
            getRobot().keyPress(KeyEvent.VK_3);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("5")){
            getRobot().keyPress(KeyEvent.VK_5);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("Bronce")){
            getRobot().keyPress(KeyEvent.VK_B);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("Sin Datos")){
            getRobot().keyPress(KeyEvent.VK_S);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("oro")){
            getRobot().keyPress(KeyEvent.VK_O);
            getRobot().keyPress(KeyEvent.VK_ENTER);
            getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
    }

    public void PressOnActivar(){
         getRobot().mouseMove(878 , 756);
         getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
         getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void FixError() throws Exception{
        Cell ErrorCell;
        for (Row row : DataExtraction.GetResultSheet()) {
            ErrorCell = row.getCell(1);
            if (ErrorCell.toString().equals("No Preactivada")) {
                LineaError = row.getCell(0).getStringCellValue();
                AccountNumber = DataExtraction.GetCuenta();
                EnterToAccount();
                PressOnVPN();
                GetGrupoDeUsuario();
                CopyTheGU();
                SearchGU();
                CopyCuentasDeGrupoUsuario();
                SearchAndSelectCuenta();
                SelectPortabilidad();
                IntroduceTheLineNumber();
                IntroduceExtension();
                PressOnContinuar();
                Thread.sleep(1000);
                if(SysScreen.has(ErrorReservaImage)){
                    FixErrorReserva();
                }
//                IntroduceSfid();
//                IntroduceTarjetaSim();
//                SelectRoaming();
//                SelectCatalogo();
//                SelectPlan();
//                SelectPerfilDeConsumo();
//                PressOnActivar();
//                CleanScreen();
            }
        }
    }
}
