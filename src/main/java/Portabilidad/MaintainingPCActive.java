package Portabilidad;

import java.awt.*;
import java.awt.event.InputEvent;
import java.util.Calendar;
import java.util.Timer;
import java.util.TimerTask;

public class MaintainingPCActive {

    public MaintainingPCActive(){

    }
    public static void MoveMouseTill9AM(){
        Timer timer = new Timer();
        TimerTask MyTask = new TimerTask() {
            @Override
            public void run() {
                try{
                    Calendar calendar = Calendar.getInstance();
                   int hr  = calendar.get(Calendar.HOUR_OF_DAY);
                   int mm = calendar.get(Calendar.MINUTE);
                   if(hr == 9 && mm == 0){
                       System.exit(130);
                   }
                   else {
                       Robot robot = new Robot();
                       robot.mouseMove(500 , 500);
                       robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                       robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                       robot.mouseMove(600 , 600);
                       robot.mousePress(InputEvent.BUTTON1_DOWN_MASK);
                       robot.mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
                   }
                }
                catch (Exception e){
                    e.getCause();
                }
            }

        };
        timer.schedule(MyTask , 0 , 6000);
    }
}
