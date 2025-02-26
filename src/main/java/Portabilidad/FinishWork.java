package Portabilidad;

import java.awt.*;
import java.awt.event.KeyEvent;

public class FinishWork {
   private static  Robot Bot;
static {
    try{
        Bot = new Robot();
    }
    catch (AWTException AW){
        AW.getCause();
    }
}
    public static void Finish() throws Exception{
        for(int i = 0; i < 20; i++) {
            Bot.keyPress(KeyEvent.VK_ALT);
            Bot.keyPress(KeyEvent.VK_ESCAPE);
            Bot.keyRelease(KeyEvent.VK_ALT);
            Bot.keyRelease(KeyEvent.VK_ESCAPE);

            Bot.keyPress(KeyEvent.VK_WINDOWS);
            Bot.keyPress(KeyEvent.VK_UP);
            Bot.keyRelease(KeyEvent.VK_WINDOWS);
            Bot.keyRelease(KeyEvent.VK_UP);

            Bot.keyPress(KeyEvent.VK_ALT);
            Bot.keyPress(KeyEvent.VK_F4);
            Bot.keyRelease(KeyEvent.VK_ALT);
            Bot.keyRelease(KeyEvent.VK_F4);

            Thread.sleep(50);
            Bot.keyPress(KeyEvent.VK_ENTER);
            Bot.keyRelease(KeyEvent.VK_ENTER);

            Thread.sleep(500);
        }
    }
}
