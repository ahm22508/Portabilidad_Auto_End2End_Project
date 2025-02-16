package TestEnviroment;


import Portabilidad.DataExtraction;
import Portabilidad.Tarea_136;

public class Main {
    public static void main(String[] args) throws Exception{
//        DataExtraction.GetOperationFile();
//        DataExtraction extract = new DataExtraction();
//                    DataExtraction.GetOperationFile();
//                    DataExtraction.GetResultSheet();
//                    extract.GetPlanesSheet();
//                    extract.ExtractSfid();
//                    extract.ExtractNumeroMovil();
        Tarea_136 tarea136 = new Tarea_136();
        DataExtraction.GetOperationFile();
        tarea136.executeTarea_135methods();
    }
}


