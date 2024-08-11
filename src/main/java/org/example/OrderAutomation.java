package org.example;

import com.sun.jna.platform.win32.WinDef;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class OrderAutomation {

    public static void main(String[] args) throws IOException, AWTException {
        // Проверка и переключение раскладки клавиатуры на английскую
        System.out.println("Проверка текущей раскладки клавиатуры.");

        KeyboardLayoutManager.ensureEnglishLayout();

        // Путь к файлу Excel
        String excelFilePath = "C:\\Users\\andru\\Documents\\AB-Nummer.xlsx";
        System.out.println("Открытие файла Excel: " + excelFilePath);

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Получаем первый лист из файла Excel
            Sheet sheet = workbook.getSheetAt(0);

            // Название окна, с которым будет работать программа
            String windowTitle = "Lutz1 MMB (Standard.zoc)";
            System.out.println("Поиск окна ZOC с заголовком: " + windowTitle);
            WinDef.HWND hwnd = MyUser32.INSTANCE.FindWindowA(null, windowTitle);

            if (hwnd == null) {
                System.out.println("Окно ZOC не найдено! Завершение работы программы.");
                return;
            }

            // Активация найденного окна ZOC
            System.out.println("Окно ZOC найдено. Активация окна.");
            Robot robot = new Robot();
            MyUser32.INSTANCE.SetForegroundWindow(hwnd);

            // Создание экземпляра OrderProcessor для обработки строк из Excel
            OrderProcessor orderProcessor = new OrderProcessor(robot, windowTitle);

            // Обработка каждой строки в листе Excel
            for (Row row : sheet) {
                orderProcessor.processOrder(row);
            }
        }

        System.out.println("Все заказы обработаны. Закрытие файла Excel.");
    }
}
