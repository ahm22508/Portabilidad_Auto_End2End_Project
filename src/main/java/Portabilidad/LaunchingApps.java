package Portabilidad;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.sikuli.script.Screen;

import java.awt.*;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;


public class LaunchingApps {
  private final WebDriver driver = new EdgeDriver();
  private final WebDriverWait wait = new WebDriverWait(driver , Duration.ofSeconds(100));
  private final File CredentialsFile= new File("C:\\Portabilidad_Auto_End2End\\data\\Credentials.xlsx");
  private final FileInputStream OpenFile = new FileInputStream(CredentialsFile);
  private final Workbook CredentialsCitrix = new XSSFWorkbook(OpenFile);
  private final Sheet Credentials = CredentialsCitrix.getSheet("CredentialOfCitrix");
  private final Screen PCscreen = new Screen();
  private final robot Bot = robot.getBotInstance();

  public LaunchingApps() throws Exception {
  }

    public String GetUserName(){
    String UserName = "";
        for(Row row : Credentials){
         UserName= row.getCell(0).getStringCellValue();
         break;
        }
        return UserName;
    }
    public String GetPassword(){
      String Password = "";
      for(Row row : Credentials){
        Password= row.getCell(1).getStringCellValue();
        break;
      }
      return Password;
    }
    public void OpenCitrix(){
      driver.navigate().to("https://essh.caas-int.vodafone.com/Citrix/ESSHWeb/");
    }
    public void Login(){
      wait.until(ExpectedConditions.elementToBeClickable(By.id("login")));
      driver.findElement(By.id("login")).sendKeys(GetUserName());
      wait.until(ExpectedConditions.elementToBeClickable(By.id("passwd")));
      driver.findElement(By.id("passwd")).sendKeys(GetPassword());
    }
    public void AcceptConditions(){
      wait.until(ExpectedConditions.elementToBeClickable(By.id("nsg-eula-accept")));
      driver.findElement(By.id("nsg-eula-accept")).click();
    }
    public void ClickOnSubmit(){
      wait.until(ExpectedConditions.elementToBeClickable(By.id("Logon")));
      driver.findElement(By.id("Logon")).click();
    }
    public void ClickOnMyHomeButton(){
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class = 'embedded-home-svg headerButton']")));
        driver.findElement(By.xpath("//div[@class = 'embedded-home-svg headerButton']")).click();
    }
    public void ClickOneEdgeChromium(){
      wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt ='Edge Chromium']")));
      driver.findElement(By.xpath("//img[@alt ='Edge Chromium']")).click();
    }
  public void ClickOnClarify115(){
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt ='Clarify 115']")));
    driver.findElement(By.xpath("//img[@alt ='Clarify 115']")).click();
  }   public void ClickOnClarify101(){
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt ='Clarify 101']")));
    driver.findElement(By.xpath("//img[@alt ='Clarify 101']")).click();
  }

    public void ClickOneWindowsExplorer(){
        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//img[@alt ='Windows Explorer']")));
        driver.findElement(By.xpath("//img[@alt ='Windows Explorer']")).click();
    }
    public void LoginToGescore() throws Exception{
      Thread.sleep(1500);
      Bot.getRobot().keyPress(KeyEvent.VK_CONTROL);
      Bot.getRobot().keyPress(KeyEvent.VK_T);
      Bot.getRobot().keyRelease(KeyEvent.VK_CONTROL);
      Bot.getRobot().keyRelease(KeyEvent.VK_T);
      Thread.sleep(1000);
      Bot.getRobot().mouseMove(275, 50);
      Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
      Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
      Thread.sleep(900);
      PCscreen.type("https://10.167.218.11/gescoresuitedashboard/");
      Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
      Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);
      Thread.sleep(10000);
      Bot.getRobot().mouseMove(380, 326);
      Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
      Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
      for (char Letter : GetUserName().toCharArray()) {
        Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(Letter));
      }
      Thread.sleep(200);
      Bot.getRobot().mouseMove(385, 396);
      Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
      Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
      PCscreen.type(GetPassword());
      Bot.getRobot().mouseMove(354, 445);
      Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
      Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
    }
    public void OpenGescore() throws Exception {
          Thread.sleep(90000);
      if (PCscreen.has("C:\\Portabilidad_Auto_End2End\\img\\MaxBrowser.png")) {
        PCscreen.click("C:\\Portabilidad_Auto_End2End\\img\\MaxBrowser.png");
        LoginToGescore();
        return;
      }
        if(PCscreen.has("C:\\Portabilidad_Auto_End2End\\img\\MaxBrowser1.png")) {
          LoginToGescore();
        }
    }
    public void RemoveGmmsAndShorCuts()throws Exception{
      PCscreen.wait("C:\\Portabilidad_Auto_End2End\\img\\CloseShortCut.png" , 300);
      PCscreen.click("C:\\Portabilidad_Auto_End2End\\img\\CloseShortCut.png");
      PCscreen.wait("C:\\Portabilidad_Auto_End2End\\img\\CloseGmms.png" , 100);
      PCscreen.click("C:\\Portabilidad_Auto_End2End\\img\\CloseGmms.png");
    }
    public void CleanScreen(){
      Bot.getRobot().keyPress(KeyEvent.VK_WINDOWS);
      Bot.getRobot().keyPress(KeyEvent.VK_M);
      Bot.getRobot().keyRelease(KeyEvent.VK_WINDOWS);
      Bot.getRobot().keyRelease(KeyEvent.VK_M);
    }

    public void ExecuteLaunchingAppMethods() throws Exception{
    OpenCitrix();
    Login();
    AcceptConditions();
    ClickOnSubmit();
    ClickOnMyHomeButton();
    ClickOneEdgeChromium();
    OpenGescore();
    ClickOneWindowsExplorer();
    Thread.sleep(30000);
    ClickOnClarify101();
    Thread.sleep(30000);
    ClickOnClarify115();
    Thread.sleep(120000);
    RemoveGmmsAndShorCuts();
    Thread.sleep(30000);
    CleanScreen();
    driver.quit();
    }

}
