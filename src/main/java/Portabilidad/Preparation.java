package Portabilidad;


import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;

public class Preparation {

    private final robot Bot = robot.getBotInstance();

    private final Windows clarify = new Windows();

    public Preparation() throws Exception {
    }

    public void prepareSystem() throws  Exception{
        int i = 3;
        String FakeLine = "666666666";
        while (i > 0) {
            //Show the system
            clarify.ShowWindow("Clarify");
            //Press on Select
            Thread.sleep(500);
            Bot.getRobot().mouseMove(226, 37);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            //press on Servicios
            Thread.sleep(500);
            Bot.getRobot().mouseMove(246, 654);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            //press on Identificador de Servicio
            Bot.getRobot().mouseMove(1183, 254);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            Thread.sleep(500);
            for(char digit : FakeLine.toCharArray()){
                Bot.getRobot().keyPress(KeyEvent.getExtendedKeyCodeForChar(digit));
            }
            //Press on Buscar
            Bot.getRobot().mouseMove(1476, 464);
            Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            //Press on OK
            Thread.sleep(500);
            Bot.getRobot().mouseMove(764, 475);
            Bot.getRobot().mousePress(KeyEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(KeyEvent.BUTTON1_DOWN_MASK);
            //Press on Cerrar Button.
            Bot.getRobot().mouseMove(977, 764);
            Bot.getRobot().mousePress(InputEvent.BUTTON1_DOWN_MASK);
            Bot.getRobot().mouseRelease(InputEvent.BUTTON1_DOWN_MASK);
            i--;
        }
    }
}