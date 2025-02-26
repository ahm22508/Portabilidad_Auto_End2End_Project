package TestEnviroment;


import Portabilidad.FinishWork;
import Portabilidad.LaunchingApps;

public class Main {
    public static void main(String[] args) throws Exception {
        LaunchingApps launchingApps = new LaunchingApps();
      launchingApps.ExecuteLaunchingAppMethods();
        FinishWork.Finish();
    }
}





