package Portabilidad;


import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.AnchorPane;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.stage.Stage;

import java.util.concurrent.Executor;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;


@SuppressWarnings("ALL")
public class Main extends Application {
    @Override
    public void start(Stage frame) {

        frame = new Stage();
        AnchorPane Container = new AnchorPane();

        Image BackGround = new Image("file:///C:/Portabilidad_Auto_End2End/data/BackGround.PNG");
        ImageView ShowImage = new ImageView(BackGround);
        ShowImage.setFitHeight(400);
        ShowImage.setFitWidth(400);
        ShowImage.setPreserveRatio(true);
        Container.getChildren().add(ShowImage);

        //Add Field.
        TextField IDCaso = new TextField();
        AnchorPane.setBottomAnchor(IDCaso, 110.0);
        AnchorPane.setLeftAnchor(IDCaso, 110.0);
        Container.getChildren().add(IDCaso);

        // Add Label to PA.
        Label labelPA = new Label("Chequear PA");
        AnchorPane.setLeftAnchor(labelPA, 20.0);
        AnchorPane.setTopAnchor(labelPA, 110.0);
        labelPA.setFont(new Font("Arial", 14));
        labelPA.setTextFill(Color.BLACK);
        labelPA.setStyle("-fx-font-weight: bold;");
        Container.getChildren().add(labelPA);

        Label labelAC = new Label("Chequear AC");
        labelAC.setFont(new Font("Arial", 14));
        labelAC.setTextFill(Color.BLACK);
        labelAC.setStyle("-fx-font-weight: bold;");
        AnchorPane.setRightAnchor(labelAC, 50.0);
        AnchorPane.setTopAnchor(labelAC, 110.0);
        Container.getChildren().add(labelAC);

        //Add PA Button
        Button ButtonPA = new Button("PA");
        ButtonPA.setOnAction(event -> {
            try {
                GetFile getFile = new GetFile();
                Preparation prepare = new Preparation();
                Check check = new Check();
                FileHandling Handle = new FileHandling();
                DataExtraction extract = new DataExtraction();
                Pre_activation preactivar = new Pre_activation();
                FileLogging fileLog = new FileLogging();

               //  getFile.ExecuteGetFileMethods();
               //  Handle.ExecuteFileHandlingMethods();

             Thread FirstThread = new Thread(()->{
                    try {

                      prepare.prepareSystem();
                    }
                    catch (Exception e){
                        e.printStackTrace();
                    }
                });
                Thread SecondThread = new Thread(()->{
                   try {
                       extract.GetOperationFile();
                       extract.GetResultSheet();
                       extract.GetPlanesSheet();
                       extract.ExtractNumeroMovil();
                   }
                   catch (Exception e){
                       e.printStackTrace();
                   }
                });
                FirstThread.start();
                SecondThread.start();
                FirstThread.join();
                SecondThread.join();
                check.ExecuteCheckPAMethods();
                preactivar.FixError();
                fileLog.ExecuteFileLoggingMethods();

            } catch (Exception e) {
                System.out.println("Exception Occured: " + e.getMessage().toString());
            }
        });
        AnchorPane.setRightAnchor(ButtonPA, 320.0);
        AnchorPane.setTopAnchor(ButtonPA, 155.0);
        Container.getChildren().add(ButtonPA);

        //Add AC Button

        Button ButtonAC = new Button("AC");
        AnchorPane.setTopAnchor(ButtonAC, 155.0);
        AnchorPane.setLeftAnchor(ButtonAC, 290.0);
        ButtonAC.setOnAction(event -> {
            try {
                GetFile getFile = new GetFile();
                Preparation prepare = new Preparation();
                Check check = new Check();
                FileHandling Handle = new FileHandling();
                DataExtraction extract = new DataExtraction();
                Pre_activation preactivar = new Pre_activation();
                FileLogging fileLog = new FileLogging();
                  getFile.ExecuteGetFileMethods();
                  Handle.ExecuteFileHandlingMethods();
                 Thread FirstThread = new Thread(()->{
                      try {
                          extract.GetOperationFile();
                          extract.GetResultSheet();
                          extract.ExtractNumeroMovil();
                      }
                      catch (Exception e){
                          e.getCause();
                      }
                  });
                  Thread SecondThread = new Thread(()->{
                      try {

                          prepare.prepareSystem();
                      }
                      catch (Exception e){
                          e.getCause();
                      }
                  });
                SecondThread.start();
                FirstThread.start();
                SecondThread.join();
                FirstThread.join();
                check.ExecuteCheckACMethods();
                fileLog.ExecuteFileLoggingMethods();
            } catch (Exception E) {
                System.out.println("Exception Occured: " + E.getMessage().toString());
            }
        });
        Container.getChildren().add(ButtonAC);
        //Show the Frame
        Scene scen = new Scene(Container, 400, 350, Color.TRANSPARENT);
        frame.setResizable(false);
        frame.setMaximized(false);
        frame.setTitle("Chequear_Lineas_Portadas");
        frame.setScene(scen);
        frame.show();
    }
    public static void main(String[] args) throws Exception {
           launch(args);
    }
}