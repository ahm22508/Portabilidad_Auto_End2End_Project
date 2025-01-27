package NewFunctionality;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.util.LinkedList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


public class DataExtraction {

    private final Sheet RevisionSheet;
    private final Sheet PlanesSheet;
    private final LinkedList<LinkedList<String>> AllData = new LinkedList<>();
    private  Cell SfidCell = null;

    public LinkedList<LinkedList<String>> GetAllData() {
        return AllData;
    }

    public Cell GetSfidCell(){
        return SfidCell;
    }

    public DataExtraction() throws Exception {
        File RevisionFile = new File("C:\\Portabilidad_Auto_End2End\\data\\operation.xlsx");
        try (FileInputStream EnterToFile = new FileInputStream(RevisionFile);
             Workbook operationSheet = new XSSFWorkbook(EnterToFile)) {
            RevisionSheet = operationSheet.getSheet("Table1");
            File PlanesFile = new File("C:\\Portabilidad_Auto_End2End\\data\\PlanGsm.xlsx");
            FileInputStream ReadPlanes = new FileInputStream(PlanesFile);
            Workbook PlanesWorkBook = new XSSFWorkbook(ReadPlanes);
            PlanesSheet = PlanesWorkBook.getSheet("Planes");
        }
    }

    
    public void ExtractAllData() {
        LinkedList<String> NumeroMovil = new LinkedList<>();
        LinkedList<String> NumeroDeCuenta = new LinkedList<>();
        LinkedList<String> NumeroDeTarjeta = new LinkedList<>();
        LinkedList<String> PerfilConsumo = new LinkedList<>();
        LinkedList <String> TarifaDeGsm = new LinkedList<>();
        LinkedList<String> CatalogoDeGsm = new LinkedList<>();
        LinkedList<String> Extension = new LinkedList<>();
        Cell LineaCell = null;
        Cell PortabilidadCell;
        Pattern TarjetaSim = Pattern.compile("(\\d{12})(\\d)");

        Pattern PerfilPattern = Pattern.compile("300mb|300MB|^\\dGB$|^\\d{2}GB$|^\\d{2}$|^\\d$|Plata|PLATA|300|Bronce|BRONCE|ORO|oro|bronce|Oro|plata|SIN_DATOS|sin_datos|Sin_Datos|sin datos| SIN DATOS| Sin Datos");
        Pattern ExtPattern = Pattern.compile("^\\d+$");
        Matcher ExtMatch;
        Matcher MatchTarjeta;
        Matcher MatchPerfilDeConsumo;
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.toString().matches("^7.?\\d{8}|^6.?\\d{8}")) {
                    LineaCell = row.getCell(cell.getColumnIndex());
                    break;
                }
            }

            if (LineaCell != null) {
                outerLoop:
                for (Row ZeroRow : RevisionSheet) {
                    for (Cell ZeroCell : ZeroRow) {
                        if (ZeroCell.getRowIndex() >= 1) {
                            if (ZeroCell.toString().contains("portabilidad") || ZeroCell.toString().contains("Portabilidad")) {
                                PortabilidadCell = ZeroCell;
                                for (Cell TargetCell : ZeroRow) {
                                    if (TargetCell.toString().contains(LineaCell.toString())) {
                                        NumeroMovil.add(LineaCell.toString());
                                        for (Cell CuentaCell : ZeroRow) {
                                            if (ZeroRow.getCell(CuentaCell.getColumnIndex()) != ZeroRow.getCell(GetSfidCell().getColumnIndex()) && ZeroRow.getCell(CuentaCell.getColumnIndex()) != ZeroRow.getCell(PortabilidadCell.getColumnIndex())) {
                                                if (CuentaCell.toString().matches("^\\d{8}$")) {
                                                    NumeroDeCuenta.add(CuentaCell.toString());
                                                    for (Cell TarjetaCell : ZeroRow) {
                                                        if (ZeroRow.getCell(TarjetaCell.getColumnIndex()) != ZeroRow.getCell(LineaCell.getColumnIndex()) && ZeroRow.getCell(TarjetaCell.getColumnIndex()) != ZeroRow.getCell(CuentaCell.getColumnIndex()) && ZeroRow.getCell(TarjetaCell.getColumnIndex()) != ZeroRow.getCell(GetSfidCell().getColumnIndex()) && !ZeroRow.getCell(TarjetaCell.getColumnIndex()).toString().contains("portabilidad") && !ZeroRow.getCell(TarjetaCell.getColumnIndex()).toString().contains("Portabilidad")) {
                                                            MatchTarjeta = TarjetaSim.matcher(TarjetaCell.toString());
                                                            if (MatchTarjeta.find()) {
                                                                String DigitoDeControl = MatchTarjeta.group(2);
                                                                NumeroDeTarjeta.add(MatchTarjeta.group());
                                                                for (Cell PerfilDeConsumo : ZeroRow) {
                                                                    if (ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()) != ZeroRow.getCell(TarjetaCell.getColumnIndex()) && ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()) != ZeroRow.getCell(LineaCell.getColumnIndex()) && ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()) != ZeroRow.getCell(GetSfidCell().getColumnIndex()) && !ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()).toString().contains("portabilidad") && !ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()).toString().contains("Portabilidad") && ZeroRow.getCell(PerfilDeConsumo.getColumnIndex()) != ZeroRow.getCell(CuentaCell.getColumnIndex())) {
                                                                        MatchPerfilDeConsumo = PerfilPattern.matcher(PerfilDeConsumo.toString());
                                                                        if (MatchPerfilDeConsumo.find()) {
                                                                            PerfilConsumo.add(MatchPerfilDeConsumo.group());
                                                                            for (Cell TarifaCell : ZeroRow) {
                                                                                if (ZeroRow.getCell(TarifaCell.getColumnIndex()) != ZeroRow.getCell(LineaCell.getColumnIndex()) && ZeroRow.getCell(TarifaCell.getColumnIndex()) != ZeroRow.getCell(CuentaCell.getColumnIndex()) && ZeroRow.getCell(TarifaCell.getColumnIndex()) != ZeroRow.getCell(GetSfidCell().getColumnIndex()) && !ZeroRow.getCell(TarifaCell.getColumnIndex()).toString().contains("portabilidad") && !ZeroRow.getCell(TarifaCell.getColumnIndex()).toString().contains("Portabilidad") && ZeroRow.getCell(TarifaCell.getColumnIndex()) != ZeroRow.getCell(TarjetaCell.getColumnIndex()) && ZeroRow.getCell(TarifaCell.getColumnIndex()) != ZeroRow.getCell(PerfilDeConsumo.getColumnIndex())) {
                                                                                    outerLoopPlanSheet:
                                                                                    for (Row Planrow : PlanesSheet) {
                                                                                        for (Cell GsmCell : Planrow) {
                                                                                            if (GsmCell.toString().equals(TarifaCell.toString())) {
                                                                                                CatalogoDeGsm.add(Planrow.getCell(1).toString());
                                                                                                TarifaDeGsm.add(TarifaCell.toString());
                                                                                                for (Cell ExtCell : ZeroRow) {
                                                                                                    if (ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(TarifaCell.getColumnIndex()) && ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(LineaCell.getColumnIndex()) && ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(CuentaCell.getColumnIndex()) && ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(GetSfidCell().getColumnIndex()) && !ZeroRow.getCell(ExtCell.getColumnIndex()).toString().contains("portabilidad") && !ZeroRow.getCell(ExtCell.getColumnIndex()).toString().contains("Portabilidad") && ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(TarjetaCell.getColumnIndex()) && ZeroRow.getCell(ExtCell.getColumnIndex()) != ZeroRow.getCell(PerfilDeConsumo.getColumnIndex())) {
                                                                                                        ExtMatch = ExtPattern.matcher(ExtCell.toString());
                                                                                                        if (ExtMatch.find()) {
                                                                                                            Extension.add(ExtMatch.group());
                                                                                                            break outerLoopPlanSheet;
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        break outerLoop;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        AllData.add(1 ,NumeroMovil);
        AllData.add(2 , NumeroDeCuenta);
        AllData.add(3, NumeroDeTarjeta);
        AllData.add(4, PerfilConsumo);
        AllData.add(5 , TarifaDeGsm);
        AllData.add(6, CatalogoDeGsm);
        AllData.add(7, Extension);
    }

      public void ExtractSfid() {
        LinkedList<String> Sfid = new LinkedList<>();
        outerLoop:
        for (Row row : RevisionSheet) {
            for (Cell cell : row) {
                if (cell.toString().matches("(\\w?\\d+[A-Z]$)") && !cell.toString().matches("^[XYZ]\\d{7}[A-Z0-9]$")) {
                    SfidCell = row.getCell(cell.getColumnIndex());
                    Sfid.add(cell.toString());
                    break outerLoop;
                }
            }
        }
        AllData.addFirst(Sfid);
    }
    }


