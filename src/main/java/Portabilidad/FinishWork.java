package Portabilidad;

import java.awt.event.KeyEvent;

public class FinishWork {
    private static final robot Bot;

    static {
        try {
             Bot = robot.getBotInstance();
        }
        catch (Exception e){
            throw new RuntimeException(e);
        }
    }


    public static void Finish() throws Exception{
        for(int i = 0; i < 20; i++) {

            Bot.getRobot().keyPress(KeyEvent.VK_ALT);
            Bot.getRobot().keyPress(KeyEvent.VK_ESCAPE);
            Bot.getRobot().keyRelease(KeyEvent.VK_ALT);
            Bot.getRobot().keyRelease(KeyEvent.VK_ESCAPE);

            Bot.getRobot().keyPress(KeyEvent.VK_WINDOWS);
            Bot.getRobot().keyPress(KeyEvent.VK_UP);
            Bot.getRobot().keyRelease(KeyEvent.VK_WINDOWS);
            Bot.getRobot().keyRelease(KeyEvent.VK_UP);

            Bot.getRobot().keyPress(KeyEvent.VK_ALT);
            Bot.getRobot().keyPress(KeyEvent.VK_F4);
            Bot.getRobot().keyRelease(KeyEvent.VK_ALT);
            Bot.getRobot().keyRelease(KeyEvent.VK_F4);

            Thread.sleep(500);
            Bot.getRobot().keyPress(KeyEvent.VK_ENTER);
            Bot.getRobot().keyRelease(KeyEvent.VK_ENTER);

            Thread.sleep(500);
        }
    }
}
