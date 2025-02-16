package Portabilidad;





public class Main {
    public static void main(String[] args) throws Exception {
        GetFile getFile = new GetFile();
        Preparation prepare = new Preparation();
        Check check = new Check();
        FileHandling Handle = new FileHandling();
        DataExtraction extract = new DataExtraction();
        Pre_activation preactivar = new Pre_activation();
        FileLogging fileLog = new FileLogging();
        //PA Part:
               getFile.StartAppForPA();

        for (int i = 0; i < GetFile.getCaseReferPA().size(); i++) {
               getFile.EnterToSpecificCase(GetFile.getCaseReferPA().get(i));
             //  getFile.ExecuteGetFileMethods();
//                Handle.ExecuteFileHandlingMethods();
//            prepare.prepareSystem();
//            Thread ExtractThread = new Thread(()->{
//                try {
//                    extract.GetOperationFile();
//                    extract.GetResultSheet();
//                    extract.GetPlanesSheet();
//                    extract.ExtractSfid();
//                    extract.ExtractNumeroMovil();
//                }
//                catch (Exception e){
//                    e.getCause();
//                }
//        });
//                    ExtractThread.start();
//                    ExtractThread.join();
//                  check.ExecuteCheckPAMethods();
//                  preactivar.FixError();

//                fileLog.ExecuteFileLoggingMethods();
      //  }

        //AC Part;

      //  getFile.StartAppForAC();

     //   for (int i = 0; i < GetFile.getCaseReferAC().size(); i++) {
       //     getFile.EnterToSpecificCase(GetFile.getCaseReferAC().get(i));
//            getFile.ExecuteGetFileMethods();
//            Handle.ExecuteFileHandlingMethods();666666
//            prepare.prepareSystem();
//            extract.GetOperationFile();
//            extract.GetResultSheet();
//            extract.GetPlanesSheet();
//            extract.ExtractNumeroMovil();
//            check.ExecuteCheckACMethods();
//            fileLog.ExecuteFileLoggingMethods();
        }
    }
}