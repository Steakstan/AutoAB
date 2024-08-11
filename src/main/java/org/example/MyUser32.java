package org.example;

import com.sun.jna.Library;
import com.sun.jna.Native;
import com.sun.jna.platform.win32.WinDef;

public interface MyUser32 extends Library {
    MyUser32 INSTANCE = Native.load("user32", MyUser32.class);

    WinDef.HWND FindWindowA(String lpClassName, String lpWindowName);
    void SetForegroundWindow(WinDef.HWND hWnd);
    WinDef.HKL GetKeyboardLayout(int idThread);
    int GetWindowThreadProcessId(WinDef.HWND hWnd, WinDef.DWORDByReference lpdwProcessId);

    // Добавление метода для получения активного окна
    WinDef.HWND GetForegroundWindow();
}
