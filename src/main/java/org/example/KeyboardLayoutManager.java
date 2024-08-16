package org.example;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.WinDef;

import java.awt.*;
import java.awt.event.KeyEvent;

public class KeyboardLayoutManager {

    private static final Object lock = new Object();
    private static boolean isEnglishLayout = false;

    public static void ensureEnglishLayout() throws AWTException, InterruptedException {
        Robot robot = new Robot();
        WinDef.HWND hwnd = MyUser32.INSTANCE.GetForegroundWindow();

        synchronized (lock) {
            while (!isEnglishLayout) {
                String currentLayoutId = getCurrentLayoutId(hwnd);
                if (isCurrentLayoutEnglish(currentLayoutId)) {
                    isEnglishLayout = true;
                    lock.notifyAll(); // Уведомляем все ожидающие потоки
                } else {
                    System.out.println("Текущая раскладка не является английской. Переключение раскладки.");
                    pressAltShift(robot);
                    lock.wait(1000); // Ждем 1 секунду
                }
            }
        }
        System.out.println("Раскладка успешно переключена на английскую.");
    }

    private static String getCurrentLayoutId(WinDef.HWND hwnd) {
        int threadId = MyUser32.INSTANCE.GetWindowThreadProcessId(hwnd, null);
        WinDef.HKL currentLayout = MyUser32.INSTANCE.GetKeyboardLayout(threadId);
        long currentLayoutValue = Pointer.nativeValue(currentLayout.getPointer());
        return String.format("%08X", currentLayoutValue);
    }

    private static boolean isCurrentLayoutEnglish(String layoutId) {
        System.out.println("Текущая раскладка: " + layoutId);
        return layoutId.startsWith("00000409") || layoutId.equals("04090409") ||
                layoutId.equals("04090809") || layoutId.equals("08090809");
    }

    private static void pressAltShift(Robot robot) {
        robot.keyPress(KeyEvent.VK_ALT);
        robot.keyPress(KeyEvent.VK_SHIFT);
        robot.keyRelease(KeyEvent.VK_SHIFT);
        robot.keyRelease(KeyEvent.VK_ALT);
    }
}
