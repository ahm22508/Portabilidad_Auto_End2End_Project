package Portabilidad;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;



public class DataExtraction {

    private static Sheet RevisionSheet;
    private static Sheet ResultSheet;
    private static Sheet PlanesSheet;
    private static Cell SfidCell;
    private static String Sfid;
    private static String Plan;
    private static Cell cellLine;
    private static Cell cellCuenta;
    private static Cell cellExtension;
    private static Cell cellTarjeta;
    private static Cell cellPlan;
    private static Workbook operationSheet;
    private static File RevisionFile;

    public DataExtraction(){

    }

    public void GetPlanesSheet() throws Exception {
            File PlanesFile = new File("C:\\Portabilidad_Auto_End2End\\data\\PlanGsm.xlsx");
            FileInputStream ReadPlanes = new FileInputStream(PlanesFile);
            Workbook PlanesWorkBook = new XSSFWorkbook(ReadPlanes);
            PlanesSheet = PlanesWorkBook.getSheet("Planes");
        }

    public static Sheet GetResultSheet(){
        if(operationSheet.getSheet("Result") != null){
            ResultSheet = operationSheet.getSheet("Result");
        }
        else{
            ResultSheet = operationSheet.createSheet("Result");
        }
        return ResultSheet;
    }

    public static void GetOperationFile() throws Exception {
         RevisionFile = new File("C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx");
        FileInputStream EnterToFile = new FileInputStream(RevisionFile);
              operationSheet = new XSSFWorkbook(EnterToFile);
            RevisionSheet = operationSheet.getSheet("Table1");
        }

        public static Workbook getWorkbook(){
        return operationSheet;
        }
        public static File getFile(){
        return RevisionFile;
        }

    public void ExtractNumeroMovil(){
        ExtractSfid();
        int RowNum = 0;
        for (Row row : RevisionSheet) {
            OuterLoop:
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1 && cell.getColumnIndex() != SfidCell.getColumnIndex()) {
                    if (cell.toString().contains("portabilidad") || cell.toString().contains("Portabilidad")) {
                        for (Cell TargetCell : row) {
                            if (TargetCell.toString().matches("^7.?\\d{8}|^6.?\\d{8}")) {
                                ResultSheet.createRow(RowNum++).createCell(0).setCellValue(TargetCell.toString());
                                cellLine = TargetCell;
                                break OuterLoop;
                            }
                        }
                    }
                }
            }
        }
    }

    public void ExtractSfid() {
        outerLoop:
        for (Row row : RevisionSheet) {
            for (Cell sfidCell : row) {
                if (row.getCell(sfidCell.getColumnIndex()).getRowIndex() >= 1) {
                    if ((!sfidCell.toString().contains("Portabilidad") || !sfidCell.toString().contains("portabilidad"))) {
                        if (sfidCell.toString().matches("([^YyZz]\\w?\\d+[A-Z]$)")) {
                            SfidCell = row.getCell(sfidCell.getColumnIndex());
                            Sfid = sfidCell.getStringCellValue();
                            break outerLoop;
                        }
                    }
                }
            }
        }
    }


    public static String GetCuenta() {
        String NumeroDeCuenta = "";
        OuterLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1) {
                    if (cell.toString().equals(Pre_activation.getLineaError()) ) {
                        for (Cell CuentaCell : row) {
                            if(CuentaCell.getColumnIndex() != SfidCell.getColumnIndex() && CuentaCell.getColumnIndex() != cellLine.getColumnIndex() && (!CuentaCell.toString().contains("Portabilidad") || !CuentaCell.toString().contains("portabilidad")))
                                if (CuentaCell.toString().matches("\\d{8}")) {
                                    cellCuenta = CuentaCell;
                                    NumeroDeCuenta = CuentaCell.toString();
                                    break OuterLoop;
                                }
                            }
                        }
                    }
                }
            }
        return NumeroDeCuenta;
    }
    public static String ExtractExtension(){
        extractTarjetaSIM();
        String Extension = "";
        OuterLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1) {
                    if (cell.toString().equals(Pre_activation.getLineaError())) {
                        for (Cell ExtensionCell : row) {
                            if (ExtensionCell.getColumnIndex() != SfidCell.getColumnIndex() && ExtensionCell.getColumnIndex() != cellLine.getColumnIndex() && ExtensionCell.getColumnIndex() != cellCuenta.getColumnIndex() && ExtensionCell.getColumnIndex() != cellTarjeta.getColumnIndex()&& (!ExtensionCell.toString().contains("Portabilidad") || !ExtensionCell.toString().contains("portabilidad"))) {
                                if (ExtensionCell.toString().matches("\\d+")){
                                    cellExtension = ExtensionCell;
                                    Extension = ExtensionCell.toString();
                                    break OuterLoop;
                                }
                            }
                        }
                    }
                }
            }
        }
        return Extension;
    }
    public static String extractTarjetaSIM(){
        String TarjetaSim = "";
        OuterLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1) {
                    if (cell.toString().equals(Pre_activation.getLineaError())) {
                        for (Cell TarjetaSimCell : row) {
                            if (TarjetaSimCell.getColumnIndex() != SfidCell.getColumnIndex() && TarjetaSimCell.getColumnIndex() != cellLine.getColumnIndex() && TarjetaSimCell.getColumnIndex() != cellCuenta.getColumnIndex() && (!TarjetaSimCell.toString().contains("Portabilidad") || !TarjetaSimCell.toString().contains("portabilidad"))) {
                                if (TarjetaSimCell.toString().matches("^\\d{13}$")){
                                    TarjetaSim = TarjetaSimCell.toString();
                                    cellTarjeta = TarjetaSimCell;
                                    break OuterLoop;
                                }
                                if(TarjetaSimCell.toString().matches("^3456\\d{13}")){
                                    TarjetaSim = TarjetaSimCell.toString().replace("3456" , "");
                                    cellTarjeta = TarjetaSimCell;
                                    break OuterLoop;
                                }
                            }
                        }
                    }
                }
            }
        }
        return TarjetaSim;
    }
    public static String getPlanDeGSM(){
        return Plan;
    }
    public static String extractCatalagoDePlan() {
        String Catalogo="";
        Cell cellInPlanSheet;
        OuterLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1) {
                    if (cell.toString().equals(Pre_activation.getLineaError())) {
                        for (Cell PlanCell : row) {
                            if (PlanCell.getColumnIndex() != cellExtension.getColumnIndex() && PlanCell.getColumnIndex() != SfidCell.getColumnIndex() && PlanCell.getColumnIndex() != cellLine.getColumnIndex() && PlanCell.getColumnIndex() != cellCuenta.getColumnIndex() && (!PlanCell.toString().contains("Portabilidad") || !PlanCell.toString().contains("portabilidad"))) {
                                for (Row rowPlan : PlanesSheet) {
                                    cellInPlanSheet = rowPlan.getCell(0);
                                    if (cellInPlanSheet != null) {
                                        if (cellInPlanSheet.toString().equals(PlanCell.toString())) {
                                            Plan = PlanCell.toString();
                                            cellPlan = PlanCell;
                                            Catalogo = rowPlan.getCell(1).getStringCellValue();
                                            break OuterLoop;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    return Catalogo;
    }

    public static String extractPerfilDeConsumo(){
        String PerfilDeConsumo = "";
        OuterLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.getRowIndex() >= 1) {
                    for (Cell PerfDeConCell : row) {
                        if (PerfDeConCell.getColumnIndex() != cellCuenta.getColumnIndex() && PerfDeConCell.getColumnIndex() != cellLine.getColumnIndex() && PerfDeConCell.getColumnIndex() != cellExtension.getColumnIndex() && PerfDeConCell.getColumnIndex() != cellTarjeta.getColumnIndex() && PerfDeConCell.getColumnIndex() != cellPlan.getColumnIndex() && PerfDeConCell.getColumnIndex() != SfidCell.getColumnIndex() && PerfDeConCell.getColumnIndex() != cellLine.getColumnIndex() && (!PerfDeConCell.toString().contains("Portabilidad") || !PerfDeConCell.toString().contains("portabilidad")))
                            if (PerfDeConCell.toString().matches("300mb|300MB|^\\dGB$|^\\d{2}GB$|^\\d{2}$|^\\d$|Plata|PLATA|300|Bronce|BRONCE|ORO|oro|bronce|Oro|plata|SIN_DATOS|sin_datos|Sin_Datos|sin datos| SIN DATOS| Sin Datos")) {
                                PerfilDeConsumo = PerfDeConCell.toString();
                                break OuterLoop;
                            }
                    }
                }
            }
        }
        return PerfilDeConsumo;
    }

    public static String GetSfid(){
        return Sfid;
    }
}