package org.example;

import com.sun.jna.platform.win32.WinDef;
import javafx.application.Application;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class OrderAutomation {

    public void processFile(int choice, String excelFilePath) throws IOException, AWTException, InterruptedException {
        System.out.println("Открытие файла Excel: " + excelFilePath);

        try (FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0);

            Row firstRow = sheet.getRow(0);
            if (firstRow == null) {
                System.out.println("Файл Excel пустой. Завершение работы программы.");
                return;
            }

            int columnCount = firstRow.getLastCellNum();
            if (choice == 1 && columnCount < 3) {
                System.out.println("Таблица не соответствует формату обработки заказов. Ожидается 3 и более колонки.");
                return;
            } else if ((choice == 2 || choice == 3) && columnCount != 2) {
                System.out.println("Таблица не соответствует формату обработки комментариев. Ожидается 2 колонки.");
                return;
            } else if (choice == 4 && columnCount != 3) {
                System.out.println("Таблица не соответствует формату обработки дат поставки. Ожидается 3 колонки.");
                return;
            }

            String windowTitle = "Lutz1 MMB Z (Standard.zoc)";
            System.out.println("Поиск окна ZOC с заголовком: " + windowTitle);
            WinDef.HWND hwnd = MyUser32.INSTANCE.FindWindowA(null, windowTitle);

            if (hwnd == null) {
                System.out.println("Окно ZOC не найдено! Завершение работы программы.");
                return;
            }

            System.out.println("Окно ZOC найдено. Активация окна.");
            Robot robot = new Robot();

            MyUser32.INSTANCE.SetForegroundWindow(hwnd);
            robot.delay(1000);  // Задержка для надежности

            // Проверяем и переключаем раскладку после активации окна
            System.out.println("Проверка текущей раскладки клавиатуры.");
            KeyboardLayoutManager.ensureEnglishLayout(hwnd);

            if (choice == 1) {
                OrderProcessor orderProcessor = new OrderProcessor(robot);
                for (Row row : sheet) {
                    orderProcessor.processOrder(row);
                }
            } else if (choice == 2) {
                CommentProcessor commentProcessor = new CommentProcessor(robot);
                for (Row row : sheet) {
                    commentProcessor.processComment(row);
                }
            } else if (choice == 3) {
                LagerbestellungProcessor lagerProcessor = new LagerbestellungProcessor(robot);
                for (Row row : sheet) {
                    lagerProcessor.processLagerComment(row);
                }
            } else if (choice == 4) {
                DeliveryDateProcessor deliveryDateProcessor = new DeliveryDateProcessor(robot);
                for (Row row : sheet) {
                    deliveryDateProcessor.processDeliveryDate(row);
                }
            }
        }

        System.out.println("Все заказы/комментарии обработаны. Закрытие файла Excel.");
    }

    public static void main(String[] args) {
        Application.launch(OrderAutomationGUIFX.class, args); // Запуск JavaFX приложения
    }
}
