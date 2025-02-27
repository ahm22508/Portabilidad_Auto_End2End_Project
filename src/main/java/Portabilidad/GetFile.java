package Portabilidad;

import org.openqa.selenium.*;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.awt.event.KeyEvent;
import java.time.Duration;
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.regex.Pattern;

public class GetFile{


    private  WebDriverWait wait;
    private  WebDriver driver;
    private int counterPA = 17;
    private int PanNumber = 4;
    private int CounterAC = 14;
    private final By IniciarButton = By.xpath("//div[@class='appmagic-button-label no-focus-outline']");
    private static final ArrayList <By> CasesReferPA = new ArrayList<>();
    private static final ArrayList <By> CasesReferAC = new ArrayList<>();
    private final robot Bot = robot.getBotInstance();

    public GetFile() throws Exception {
    }

    public static ArrayList<By> getCaseReferPA(){
        return CasesReferPA;

    }
    public static ArrayList<By> getCaseReferAC(){
        return CasesReferAC;
    }

    public void Switching() {
        wait.until(ExpectedConditions.visibilityOfElementLocated(By.tagName("iframe")));
        driver.switchTo().frame(driver.findElement(By.tagName("iframe")));
    }

    public void SwitchingOff() {
        driver.switchTo().defaultContent();
    }


    public void StartBrowser(){
        driver = new EdgeDriver();
        wait = new WebDriverWait(driver, Duration.ofSeconds(200));
    }
    public void GoToVCTool() {
        driver.navigate().to("https://apps.powerapps.com/play/e/default-68283f3b-8487-4c86-adb3-a5228f18b893/a/5996516e-15f4-4d7f-a077-65b4ec5a037f?tenantId=68283f3b-8487-4c86-adb3-a5228f18b893&hint=fb40c2e7-cffd-4e5b-9711-ec3bd2f12dca&sourcetime=1699970459236");
        driver.manage().window().maximize();
    }

    public void InitiateVCTool() {
        Switching();
        wait.until(ExpectedConditions.elementToBeClickable(IniciarButton));
        driver.findElement(IniciarButton).click();
        SwitchingOff();
    }

    public void EnterToCases() {
        Actions actions = new Actions(driver);
        actions.keyDown(Keys.TAB).build().perform();
        actions.keyDown(Keys.TAB).build().perform();
        actions.keyDown(Keys.TAB).build().perform();
        actions.keyDown(Keys.ENTER).build().perform();
    }
    public void FilterWithMovil(){
    Switching();
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='filtero por Tipo']")));
    driver.findElement(By.xpath("//span[text()='filtero por Tipo']")).click();
    wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='Móvil']")));
    driver.findElement(By.xpath("//span[text()='Móvil']")).click();
    SwitchingOff();
    }


    public void SearchPACases(){
        Switching();
        wait.until(ExpectedConditions.textMatches(By.xpath("(//div[@spellcheck='false'])["+counterPA+"]") , Pattern.compile("\\b([1-9]|1[0-2])/([1-9]|[12][0-9]|3[01])/(\\d{4})\\b")));
        String datePreactivation = driver.findElement(By.xpath("(//div[@spellcheck='false'])["+counterPA+"]")).getText();
        SwitchingOff();
        DateTimeFormatter format = DateTimeFormatter.ofPattern("M/d/yyyy");
       LocalDate DateCase = LocalDate.parse(datePreactivation , format);
        Period PreActivationperiod = Period.between(DateCase , LocalDate.now());
        if (PreActivationperiod.getDays() > 0){
            return;
        }
        if (PreActivationperiod.getDays() == 0) {
            CasesReferPA.add(By.xpath("(//div[@class='powerapps-icon no-focus-outline'])["+PanNumber+"]"));
            counterPA+=7;
            PanNumber+=1;
            SearchPACases();
        }

         if (PreActivationperiod.getDays() < 0) {
            counterPA+=7;
            PanNumber+=1;
           SearchPACases();
        }
    }
    public void SearchCasesAC(){
        Switching();
        wait.until(ExpectedConditions.textMatches(By.xpath("(//div[@spellcheck='false'])["+CounterAC+"]") , Pattern.compile("\\b([1-9]|1[0-2])/([1-9]|[12][0-9]|3[01])/(\\d{4})\\b")));
        String dateActivation = driver.findElement(By.xpath("(//div[@spellcheck='false'])["+CounterAC+"]")).getText();
        SwitchingOff();
        DateTimeFormatter format = DateTimeFormatter.ofPattern("M/d/yyyy");
        LocalDate DateCase = LocalDate.parse(dateActivation , format);
        Period Activationperiod = Period.between(DateCase , LocalDate.now());
        if (Activationperiod.getDays() > 0){
            return;
        }
        if (Activationperiod.getDays() == 0) {
            CasesReferAC.add(By.xpath("(//div[@class='powerapps-icon no-focus-outline'])["+PanNumber+"]"));
            CounterAC+=7;
            PanNumber +=1;
            SearchCasesAC();
        }
        if (Activationperiod.getDays() < 0) {
            CounterAC+=7;
            PanNumber +=1;
            SearchCasesAC();
        }
    }

    public  void EnterToSpecificCase(By PanIcon){
        Switching();
        driver.findElement(PanIcon).click();
        SwitchingOff();
    }


    public void DownloadFile() throws Exception {
        Thread.sleep(3500);
        Bot.getRobot().mouseMove(1331, 786);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
        Thread.sleep(1500);
        Bot.getRobot().mouseMove(441, 516);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
    }

    public void ShowFile() throws Exception {
        Thread.sleep(4500);
        Bot.getRobot().mouseMove(1302, 134);
        Thread.sleep(1000);
        Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
        Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
    }

    public void Exit() {
        driver.quit();
    }
    public void StartAppForPA() throws Exception{
        StartBrowser();
        GoToVCTool();
        InitiateVCTool();
        Thread.sleep(6000);
        EnterToCases();
        FilterWithMovil();
       // SearchPACases();
    }
    public void StartAppForAC() throws Exception {
        StartBrowser();
        GoToVCTool();
        InitiateVCTool();
        Thread.sleep(6000);
        EnterToCases();
        FilterWithMovil();
       // SearchCasesAC();
    }
    public void ExecuteGetFileMethods() throws Exception{
         DownloadFile();
         ShowFile();
         Thread.sleep(1000);
         Exit();
    }
}