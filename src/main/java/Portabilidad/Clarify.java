package Portabilidad;

import com.sun.jna.Library;
import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.win32.W32APIOptions;

public class Clarify {
    public interface User32 extends Library{
        User32 INSTANCE = Native.load("User32" , User32.class , W32APIOptions.DEFAULT_OPTIONS);
        Pointer FindWindow(String IP , String WindowTitle);
        void SetForegroundWindow(Pointer hwnd);
        void ShowWindow(Pointer hwnd, int Size);
    }
    public void ShowClarify(){
        User32 user32 = User32.INSTANCE;
        Pointer hwnd = user32.FindWindow(null , "Clarify - ClearSupport - [Conexion Clarify 11.5] - \\\\Remote");
        user32.SetForegroundWindow(hwnd);
        user32.ShowWindow(hwnd , 1);
    }
}
