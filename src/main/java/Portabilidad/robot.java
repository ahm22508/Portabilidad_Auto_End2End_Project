package Portabilidad;

import java.awt.*;

public class robot {

    private static volatile robot botInstance;
    private final Robot Bot;

    private robot() throws AWTException {
        Bot = new Robot();
    }

    public static  robot getBotInstance() throws Exception{
        if(botInstance == null){
            synchronized (robot.class){
                if(botInstance == null){
                    botInstance = new robot();
                }
            }
        }
    return botInstance;
    }
    public Robot getRobot(){
        return Bot;
    }
}

