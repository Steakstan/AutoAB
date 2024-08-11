package org.example;

import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.WinDef;

import java.awt.*;
import java.awt.event.KeyEvent;

public class KeyboardLayoutManager {

    public static void ensureEnglishLayout() throws AWTException {
        Robot robot = new Robot();
        WinDef.HWND hwnd = MyUser32.INSTANCE.GetForegroundWindow(); // Получаем текущее активное окно

        String currentLayoutId = getCurrentLayoutId(hwnd);

        // Цикл, который будет продолжаться, пока раскладка не станет английской
        while (!isCurrentLayoutEnglish(currentLayoutId)) {
            System.out.println("Текущая раскладка не является английской. Переключение раскладки.");
            pressAltShift(robot);
            try {
                // Даем немного времени системе переключить раскладку
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                System.err.println("Поток был прерван.");
            }

            // Повторная проверка после переключения раскладки
            currentLayoutId = getCurrentLayoutId(hwnd);
            System.out.println("Проверка раскладки после переключения: " + currentLayoutId);
        }

        System.out.println("Раскладка успешно переключена на английскую.");
    }

    private static String getCurrentLayoutId(WinDef.HWND hwnd) {
        int threadId = MyUser32.INSTANCE.GetWindowThreadProcessId(hwnd, null); // Получаем ID потока активного окна
        WinDef.HKL currentLayout = MyUser32.INSTANCE.GetKeyboardLayout(threadId); // Получаем раскладку для потока
        long currentLayoutValue = Pointer.nativeValue(currentLayout.getPointer());
        return String.format("%08X", currentLayoutValue);
    }

    private static boolean isCurrentLayoutEnglish(String layoutId) {
        System.out.println("Текущая раскладка: " + layoutId); // Логирование текущей раскладки

        // Проверяем, является ли текущая раскладка одной из английских
        return layoutId.startsWith("00000409") || // Английская (US)
                layoutId.equals("04090409") ||
                layoutId.equals("04090809") || // Английская (UK)
                layoutId.equals("08090809");  // Английская (другая версия)
    }

    private static void pressAltShift(Robot robot) {
        robot.keyPress(KeyEvent.VK_ALT);
        robot.keyPress(KeyEvent.VK_SHIFT);
        robot.keyRelease(KeyEvent.VK_SHIFT);
        robot.keyRelease(KeyEvent.VK_ALT);
    }
}
