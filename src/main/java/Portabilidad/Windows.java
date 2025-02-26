package Portabilidad;

import com.sun.jna.Library;
import com.sun.jna.Native;
import com.sun.jna.Pointer;

import com.sun.jna.PointerType;
import com.sun.jna.platform.win32.WinDef;
import com.sun.jna.platform.win32.WinUser;
import com.sun.jna.win32.W32APIOptions;

public class Windows {
    public interface User32 extends Library{
        User32 INSTANCE = Native.load("User32" , User32.class , W32APIOptions.DEFAULT_OPTIONS);
        Pointer FindWindow(String IP , String WindowTitle);
        void SetForegroundWindow(Pointer hwnd);
        void ShowWindow(Pointer hwnd, int Size);
        void EnumWindows(WinUser.WNDENUMPROC WindIterator , PointerType pointer);
        void GetWindowTextA(WinDef.HWND hwnd , byte [] Windows, int MaxCount);
    }
    public void ShowWindow(String mainName){
        User32 user32 = User32.INSTANCE;
        user32.EnumWindows((hwnd , pointer)->{
          byte [] windows  = new byte[512];
          user32.GetWindowTextA(hwnd , windows , 512);
          String windowTitle = Native.toString(windows);
          if (windowTitle.contains(mainName) && windowTitle.contains("Remote")){
               pointer= user32.FindWindow(null, windowTitle);
              user32.SetForegroundWindow(pointer);
              user32.ShowWindow(pointer, 1);
          }
        return true;
        }, null);

            }
    }

