package Portabilidad;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


public class Tarea_136 extends Check {

   private final File Proceso136 = new File("C:\\Portabilidad_Auto_End2End\\data\\Proceso_136.xlsx");
   private final FileInputStream openTheFile = new FileInputStream(Proceso136);
   private final Workbook ProcesoWorkBook = new XSSFWorkbook(openTheFile);

    public Tarea_136() throws Exception{
    }
    public String getCif() throws Exception{
        Clarify clarify = new Clarify();
        clarify.ShowClarify();
        Thread.sleep(1000);
        Clipboard clipboard;
       clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
        getRobot().mouseMove(217, 35);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseMove(284, 232);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().delay(1000);
        getRobot().mouseMove(466, 316);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        for(char Letter : DataExtraction.GetCuenta().toCharArray()){
            getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
        }
        getRobot().mouseMove(1466, 485);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().delay(1000);
        getRobot().mouseMove(452 , 544);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(200);
        getRobot().keyPress(KeyEvent.VK_CONTROL);
        getRobot().keyPress(KeyEvent.VK_C);
        getRobot().keyRelease(KeyEvent.VK_CONTROL);
        getRobot().keyRelease(KeyEvent.VK_C);
        Transferable Content = clipboard.getContents(null);

        getRobot().mouseMove(1481 , 760);
        getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
        getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);

         return (String)Content.getTransferData(DataFlavor.stringFlavor);
    }

    public Sheet Hoja1Sheet(){

        return ProcesoWorkBook.getSheet("Hoja1");
    }
    public void PopulateSheet(){
        long StartTime = System.currentTimeMillis();
        int i = 0;
        for (Row row : DataExtraction.GetResultSheet()) {
            Cell ResultCell = row.getCell(2);
            if (ResultCell.getStringCellValue().equals("No")) {
                i++;
                Hoja1Sheet().createRow(i).createCell(0).setCellValue("getCif()");
                Hoja1Sheet().getRow(i).createCell(1).setCellValue("Pre_activation.getLineaError()");
                Hoja1Sheet().getRow(i).createCell(2).setCellValue("24759476");
                Hoja1Sheet().getRow(i).createCell(3).setCellValue("29584847");
                Hoja1Sheet().getRow(i).createCell(4).setCellValue("4494858476575");
                Hoja1Sheet().getRow(i).createCell(5).setCellValue("Integrado");
                Hoja1Sheet().getRow(i).createCell(6).setCellValue("MPMVB");
                Hoja1Sheet().getRow(i).createCell(7).setCellValue("ORo");
                Hoja1Sheet().getRow(i).createCell(8).setCellValue("1500");
                Hoja1Sheet().getRow(i).createCell(12).setCellValue("Error Preactivacion");
                Hoja1Sheet().getRow(i).createCell(13).setCellValue("Error Preactivacion");
                Hoja1Sheet().getRow(i).createCell(14).setCellValue("Error");
                Hoja1Sheet().getRow(i).createCell(15).setCellValue("V594458A");
                long FinishTime = System.currentTimeMillis();
                System.out.println(FinishTime - StartTime);
            }
        }
    }
    public void SaveFile() throws Exception{
        FileOutputStream FileSaving  = new FileOutputStream(Proceso136);
        ProcesoWorkBook.write(FileSaving);
    }
    public void closeFile()throws Exception{
        ProcesoWorkBook.close();
    }
        public void executeTarea_135methods()throws Exception{
        PopulateSheet();
        SaveFile();
        closeFile();

        }

}
