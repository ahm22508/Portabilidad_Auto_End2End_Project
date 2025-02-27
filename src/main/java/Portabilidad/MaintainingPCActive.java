package Portabilidad;

import java.awt.event.InputEvent;
import java.util.Calendar;
import java.util.Timer;
import java.util.TimerTask;

public class MaintainingPCActive {
    private final static robot Bot;
    static {
        try {
         Bot = robot.getBotInstance();
        }
        catch (Exception e){
            throw new RuntimeException(e);
        }
    }

    public MaintainingPCActive() {

    }
    public static void MoveMouseTill9AM(){
        Timer timer = new Timer();
        TimerTask MyTask = new TimerTask() {
            @Override
            public void run() {
                    Calendar calendar = Calendar.getInstance();
                   int hr  = calendar.get(Calendar.HOUR_OF_DAY);
                   int mm = calendar.get(Calendar.MINUTE);
                   if(hr == 9 && mm == 0){
                       System.exit(130);
                   }
                   else {
                       Bot.getRobot().mouseMove(500 , 500);
                       Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                       Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                       Bot.getRobot().mouseMove(600 , 600);
                       Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
                       Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                   }
            }

        };
        timer.schedule(MyTask , 0 , 6000);
    }
}
