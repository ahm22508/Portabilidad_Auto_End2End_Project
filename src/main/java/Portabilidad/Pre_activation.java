package Portabilidad;

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



public class Pre_activation {
    private final Clipboard clipboard;
    private String GrupoUsuario;
    private String Cuentas;
    private String AccountNumber;
    private static String LineaError;
    private final Screen SysScreen = new Screen();
    private final Pattern ErrorReservaImage = new Pattern("C:\\Portabilidad_Auto_End2End\\img\\ErrorReservaLinea.png");
    private final robot Bot = robot.getBotInstance();

    public Pre_activation() throws Exception {
        clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    }

    public static String getLineaError() {
        return LineaError;
    }

    public void EnterToAccount() throws Exception {
        Thread.sleep(300);
        //Press on Select
        Bot.getRobot().mouseMove(217, 35);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Press on cuenta
        Thread.sleep(300);
        Bot.getRobot().mouseMove(246, 584);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Click On Account Field
        Thread.sleep(600);
        Bot.getRobot().mouseMove(538, 239);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Enter the Account Number
        Thread.sleep(300);
        for (char Letter : AccountNumber.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        //Press on Buscar Button
        Bot.getRobot().mouseMove(1446, 427);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Select the Account
        Thread.sleep(300);
        Bot.getRobot().mouseMove(427, 518);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Press On Abrir Button
        Bot.getRobot().mouseMove(1152, 752);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void PressOnVPN() throws Exception {
        Thread.sleep(6000);
        Bot.getRobot().mouseMove(1472, 204);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void GetGrupoDeUsuario() throws Exception {
        Thread.sleep(9000);
        //Press on Transferencia Fijo
        Bot.getRobot().mouseMove(1055, 199);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //press on AccountField
        Thread.sleep(500);
        Bot.getRobot().mouseMove(812, 293);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Enter the Account Number.
        Thread.sleep(500);
        for (char Letter : AccountNumber.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        //press on Buscar Button
        Thread.sleep(200);
        Bot.getRobot().mouseMove(918, 286);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Clean Message Error
        Thread.sleep(2000);
        Bot.getRobot().mouseMove(670, 465);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Select Grupo Usuario Result
        Thread.sleep(500);
        Bot.getRobot().mouseMove(931, 365);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(1000);
    }

    public void CopyTheGU() throws Exception {
        Transferable FirstContent = clipboard.getContents(null);
        GrupoUsuario = (String) FirstContent.getTransferData(DataFlavor.stringFlavor);
    }

    public void SearchGU() throws Exception {
        Thread.sleep(300);
        Bot.getRobot().mouseMove(811, 202);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1031, 267);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        for (char FirstLetter : GrupoUsuario.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(FirstLetter));
            break;
        }
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
    }

    public void SearchInAnotherGU() throws Exception {
        Bot.getRobot().mouseMove(1094, 272);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
        Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(600);
    }

    public void CopyCuentasDeGrupoUsuario() throws Exception {
        //Select Cuentas
        Thread.sleep(300);
        Bot.getRobot().mouseMove(610, 246);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        //Copy Cuentas
        Thread.sleep(300);
        Bot.getRobot().mousePress(InputEvent.BUTTON3_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().mouseMove(631, 263);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
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
                Bot.getRobot().mouseMove(553, 258);
                Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(50);
                for (int i = 1; i <= CuentaPosition; i++) {
                    Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
                    Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
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
        Bot.getRobot().mouseMove(724, 389);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        Bot.getRobot().mouseMove(653, 461);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void IntroduceTheLineNumber() throws Exception {
        Bot.getRobot().mouseMove(904, 393);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        for (char Letter : LineaError.toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
    }

    public void IntroduceExtension() throws Exception {
        Bot.getRobot().mouseMove(1069, 625);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        for (char Letter : DataExtraction.ExtractExtension().toCharArray()) {
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
    }

    public void PressOnContinuar() {
        Bot.getRobot().mouseMove(971, 663);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
    public void FixErrorReserva() throws Exception{
        Bot.getRobot().mouseMove(669 , 466);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(556 , 200);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(600);
        Bot.getRobot().mouseMove(726 , 320);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(778 , 743);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_V);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(886 , 735);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(15000);
        Bot.getRobot().mouseMove(62 , 137);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(100);
        Bot.getRobot().mouseMove(113 , 279);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(1286 , 487);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(2000);
        Bot.getRobot().mouseMove(355 , 274);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for(char Letter : LineaError.toCharArray()){
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_C);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_C);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(563 , 273);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyPress(KeyEvent.VK_V);
        Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
        Bot.getRobot().keyRelease(KeyEvent.VK_V);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(582 , 324);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
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
            Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
            Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
        }
        if(LinePrefix.startsWith("6")) {
            int Prefix = Integer.parseInt(LinePrefix);
            if (Prefix > 605) {
                Bot.getRobot().mouseMove(585 , 435);
                for (int coun = 0; coun < Prefix - 605; coun++) {
                    Thread.sleep(50);
                    Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                }
                Bot.getRobot().mouseMove(535 , 439);
                Thread.sleep(100);
                Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }
        }
        if(LinePrefix.startsWith("600")){
            Bot.getRobot().mouseMove(531 , 370);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("601")){
            Bot.getRobot().mouseMove(531 , 382);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("602")) {
            Bot.getRobot().mouseMove(531, 399);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("603")){
            Bot.getRobot().mouseMove(531 , 423);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("604")){
            Bot.getRobot().mouseMove(531 , 419);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LinePrefix.startsWith("605")){
            Bot.getRobot().mouseMove(531 , 435);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }

        if(LinePrefix.startsWith("711")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("717")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("722")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("727")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
        }
        if(LinePrefix.startsWith("744")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
        }

        if(LinePrefix.startsWith("747")){
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
            Bot.getRobot().keyRelease(KeyEvent.VK_7);
            Bot.getRobot().keyPress(KeyEvent.VK_7);
        }
        Bot.getRobot().mouseMove(817 , 324);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        if(LineRango.equals("56")){
            Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
            Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
        }
        if(LineRango.equals("00")){
            Bot.getRobot().mouseMove(784 , 345);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("01")){
            Bot.getRobot().mouseMove(784 , 356);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("02")){
            Bot.getRobot().mouseMove(784 , 370);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("03")){
            Bot.getRobot().mouseMove(784 , 384);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("04")){
            Bot.getRobot().mouseMove(784 , 395);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("05")){
            Bot.getRobot().mouseMove(784 , 410);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("06")){
            Bot.getRobot().mouseMove(784 , 422);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        if(LineRango.equals("07")){
            Bot.getRobot().mouseMove(784 , 435);
            Thread.sleep(100);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        int Range = Integer.parseInt(LineRango);
        if(Range > 7){
            Bot.getRobot().mouseMove(817, 437);
            for(int i =0; i < Range-7; i++){
                Thread.sleep(50);
                Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            }
            Bot.getRobot().mouseMove(793, 437);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        }
        Thread.sleep(200);
        Bot.getRobot().mouseMove(993, 324);
//        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
//        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        Bot.getRobot().mouseMove(335, 423);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1475, 597);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(1481, 760);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(1269, 760);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseMove(65, 135);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(500);
        Bot.getRobot().mouseMove(101, 264);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(600);
        Bot.getRobot().mouseMove(820, 206);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        PressOnContinuar();
    }

    public void IntroduceSfid() throws Exception {
        Thread.sleep(7000);
        String Sfid = DataExtraction.GetSfid();
        int i = 1;
        Bot.getRobot().mouseMove(320, 197);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for (char Letter : Sfid.toCharArray()) {
            i++;
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
            if (Sfid.length() - i == 1) {
                Bot.getRobot().mouseMove(426, 197);
                Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(100);
                Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Sfid.toCharArray()[i]));
            }
        }
    }

    public void CleanScreen() throws Exception {
        Bot.getRobot().mouseMove(1449, 756);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1452, 767);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1272, 758);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1454, 752);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void IntroduceTarjetaSim() throws Exception {
        int i = 0;
        Bot.getRobot().mouseMove(333, 261);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for (char Letter : DataExtraction.extractTarjetaSIM().toCharArray()) {
            i++;
            Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
            if (DataExtraction.extractTarjetaSIM().length() - i == 1) {
                Bot.getRobot().mouseMove(426, 260);
                Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                Thread.sleep(100);
                Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(DataExtraction.extractTarjetaSIM().toCharArray()[i]));
            }
        }
    }

    public void SelectRoaming() throws Exception {
        Thread.sleep(100);
        Bot.getRobot().mouseMove(1467, 382);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1320, 438);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    public void SelectCatalogo() throws Exception {
        String Catalogo = DataExtraction.extractCatalagoDePlan();

        Bot.getRobot().mouseMove(1116, 444);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        if (Catalogo.contains("Integrados")) {
            for (int i = 0; i < 6; i++) {
                Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
                Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        if (Catalogo.contains("Planes Especial Sol. Empresa VPN")) {
            for (int i = 0; i < 2; i++) {
                Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
                Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        if (Catalogo.contains("Planes Soluciones Empresa VPN")) {
            for (int i = 0; i < 4; i++) {
                Bot.getRobot().keyPress(KeyEvent.VK_DOWN);
                Bot.getRobot().keyRelease(KeyEvent.VK_DOWN);
            }
        }
        Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
        Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        Thread.sleep(200);
        Bot.getRobot().mouseMove(1210, 466);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }

    @SuppressWarnings("BusyWait")
    public void SelectPlan() throws Exception {
        Thread.sleep(4000);
        Bot.getRobot().mouseMove(869, 490);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        Bot.getRobot().mousePress(InputEvent.BUTTON3_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON3_DOWN_MASK);
        Thread.sleep(900);
        Bot.getRobot().mouseMove(900, 568);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1000);
        Transferable Planes = clipboard.getContents(null);
        String PlanesEnCatalogo = (String) Planes.getTransferData(DataFlavor.stringFlavor);


        int CoordinateX = 861;
        int CoordinateY = 494;

        java.util.regex.Pattern PlanPattern = java.util.regex.Pattern.compile("[A-Z]{5}");
        Matcher PlanMatch = PlanPattern.matcher(PlanesEnCatalogo);
        Bot.getRobot().mouseMove(861, 494);
        Thread.sleep(1000);
        while (PlanMatch.find()) {
            if (CoordinateY < 615) {
                CoordinateY += 14;
                if (PlanMatch.group().equals(DataExtraction.getPlanDeGSM())) {

                    Bot.getRobot().mouseMove(CoordinateX, CoordinateY);
                    Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                    Thread.sleep(500);
                    Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
                    Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
                    break;
                }
            }

            if (CoordinateY > 615) {
                if (PlanMatch.group().equals(DataExtraction.getPlanDeGSM())) {
                    Thread.sleep(100);
                    Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                    Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
                    Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
                    break;
                } else {
                    Thread.sleep(100);
                    Bot.getRobot().mouseMove(CoordinateX, CoordinateY);
                    Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                    Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                }
            }
        }
    }

    public void SelectPerfilDeConsumo()throws Exception{
        Thread.sleep(2500);
        Bot.getRobot().mouseMove(306 , 711);
        Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(300);
        if(DataExtraction.extractPerfilDeConsumo().contains("10")){
            Bot.getRobot().keyPress(KeyEvent.VK_1);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("1")){
            Bot.getRobot().keyPress(KeyEvent.VK_1);
            Bot.getRobot().keyPress(KeyEvent.VK_1);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("20")){
            Bot.getRobot().keyPress(KeyEvent.VK_2);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("2") || DataExtraction.extractPerfilDeConsumo().contains("plata")){
            Bot.getRobot().keyPress(KeyEvent.VK_2);
            Bot.getRobot().keyPress(KeyEvent.VK_2);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("300")){
            Bot.getRobot().keyPress(KeyEvent.VK_3);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("3")){
            Bot.getRobot().keyPress(KeyEvent.VK_3);
            Bot.getRobot().keyPress(KeyEvent.VK_3);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("5")){
            Bot.getRobot().keyPress(KeyEvent.VK_5);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("Bronce")){
            Bot.getRobot().keyPress(KeyEvent.VK_B);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("Sin Datos")){
            Bot.getRobot().keyPress(KeyEvent.VK_S);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
        if(DataExtraction.extractPerfilDeConsumo().contains("oro")){
            Bot.getRobot().keyPress(KeyEvent.VK_O);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
        }
    }

    public void PressOnActivar(){
         Bot.getRobot().mouseMove(878 , 756);
         Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
         Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
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

                //Here i will add new Code to log if the Preactivation process goes correct or not...
            }
        }
    }
}
