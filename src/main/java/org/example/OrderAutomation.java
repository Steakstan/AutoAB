package org.example;

import com.sun.jna.platform.win32.User32;
import com.sun.jna.platform.win32.WinDef;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class OrderAutomation {

    public static void main(String[] args) throws IOException, AWTException {
        String excelFilePath = "C:\\Users\\andru\\Documents\\AB-Nummer.xlsx";
        System.out.println("Открытие файла Excel: " + excelFilePath);

        // Открытие файла Excel
        FileInputStream fileInputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // Поиск и активация окна ZOC
        String windowTitle = "Lutz1 MMB (Standard.zoc)"; // Замените на точное название окна ZOC, как оно отображается в заголовке окна
        System.out.println("Поиск окна ZOC с заголовком: " + windowTitle);
        WinDef.HWND hwnd = User32.INSTANCE.FindWindow(null, windowTitle);

        if (hwnd == null) {
            System.out.println("Окно ZOC не найдено! Завершение работы программы.");
            workbook.close();
            fileInputStream.close();
            return;
        }

        System.out.println("Окно ZOC найдено. Активация окна.");
        User32.INSTANCE.SetForegroundWindow(hwnd); // Активация окна

        Robot robot = new Robot();

        for (Row row : sheet) {
            String orderNumber = getCellValueAsString(row.getCell(0));
            String positionNumber = getCellValueAsString(row.getCell(1));
            String deliveryDate = getCellValueAsString(row.getCell(3));
            String confirmationNumber = getCellValueAsString(row.getCell(2));

            System.out.println("Обработка заказа: " + orderNumber);

            // Запуск скрипта через F1
            pressF1(robot);

            // Получение значения из буфера обмена
            String clipboardContent = getClipboardContents();

            // Проверка совпадения с ожидаемым значением "311"
            while (!clipboardContent.equals("311")) {
                System.out.println("Значение позиции курсора не равно '311'. Перемещение курсора влево и повторная проверка.");
                pressLeftArrow(robot); // Нажимаем стрелку влево

                // Повторный запуск скрипта через F1
                pressF1(robot);

                // Считывание нового значения из буфера обмена
                clipboardContent = getClipboardContents();
            }

            System.out.println("Курсор на правильной позиции. Ввод номера заказа: " + orderNumber);
            typeText(robot, orderNumber);

            pressEnter(robot);
            robot.delay(1000);

            System.out.println("Ввод номера позиции: " + positionNumber);
            typeText(robot, positionNumber);
            pressEnter(robot);

            // Проверка позиции курсора после ввода номера позиции
            pressF1(robot);
            clipboardContent = getClipboardContents();

            if (clipboardContent.equals("1374")) {
                System.out.println("Курсор имеет значение '1374'. Нажатие стрелки назад и переход к следующему заказу.");
                pressLeftArrow(robot); // Возвращаемся назад
                continue; // Переход к следующей строке в таблице
            }

            System.out.println("Проверка позиции курсора перед вводом даты.");

            // Запуск скрипта через F1 для проверки позиции перед вводом даты
            pressF1(robot);

            // Получение значения из буфера обмена
            clipboardContent = getClipboardContents();

            // Проверка позиции курсора
            while (clipboardContent.equals("2480")) {
                System.out.println("Позиция курсора равна '2480'. Нажатие Enter и повторная проверка.");
                pressEnter(robot);
                pressF1(robot);
                clipboardContent = getClipboardContents();
            }

            if (clipboardContent.equals("936")) {
                System.out.println("Курсор на правильной позиции. Ввод даты поставки: " + deliveryDate);
                typeText(robot, deliveryDate);
                pressEnter(robot);
                pressEnter(robot);
            } else {
                System.out.println("Курсор не на правильной позиции для ввода даты.");
            }

            System.out.println("Ввод номера подтверждения: " + confirmationNumber);
            typeText(robot, confirmationNumber);

            // Проверка позиции курсора для подтверждения
            while (!clipboardContent.equals("2375")) {
                pressEnter(robot);
                pressF1(robot);
                clipboardContent = getClipboardContents();

                // Если значение курсора равно "411", возвращаемся к шагу ввода номера заказа
                if (clipboardContent.equals("411")) {
                    System.out.println("Курсор имеет значение '411'. Прерывание цикла и переход к следующему заказу.");
                    break; // Выходим из текущего цикла и возвращаемся к началу обработки следующего заказа
                }
            }

            // Если мы вышли из цикла из-за значения "411", переходим к следующему заказу
            if (clipboardContent.equals("411")) {
                continue; // Переход к следующей строке в таблице (начало следующей итерации цикла for)
            }

            pressEnter(robot);

            System.out.println("Задержка перед обработкой следующего заказа (2 секунды).");
            robot.delay(2000); // Задержка между итерациями
        }

        System.out.println("Все заказы обработаны. Закрытие файла Excel.");
        workbook.close();
        fileInputStream.close();
    }

    private static String getCellValueAsString(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private static void typeText(Robot robot, String text) {
        for (char c : text.toCharArray()) {
            if (Character.isUpperCase(c)) {
                robot.keyPress(KeyEvent.VK_SHIFT);  // Нажимаем Shift
            }
            int keyCode = KeyEvent.getExtendedKeyCodeForChar(Character.toUpperCase(c));
            if (KeyEvent.CHAR_UNDEFINED == keyCode) {
                throw new RuntimeException("Key code not found for character '" + c + "'");
            }
            robot.keyPress(keyCode);
            robot.keyRelease(keyCode);

            if (Character.isUpperCase(c)) {
                robot.keyRelease(KeyEvent.VK_SHIFT);  // Отпускаем Shift
            }
        }
    }

    private static void pressEnter(Robot robot) {
        System.out.println("Нажатие Enter.");
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
        robot.delay(400); // Добавляем задержку в 400 миллисекунд после каждого нажатия Enter
    }

    private static void pressF1(Robot robot) {
        robot.keyPress(KeyEvent.VK_F1);
        robot.keyRelease(KeyEvent.VK_F1);
        robot.delay(500); // Задержка, чтобы убедиться, что команда выполнена
    }

    private static void pressLeftArrow(Robot robot) {
        System.out.println("Нажатие стрелки влево.");
        robot.keyPress(KeyEvent.VK_LEFT);
        robot.keyRelease(KeyEvent.VK_LEFT);
        robot.delay(500); // Задержка после нажатия стрелки влево
    }

    private static String getClipboardContents() {
        try {
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            Transferable contents = clipboard.getContents(null);
            if (contents != null && contents.isDataFlavorSupported(DataFlavor.stringFlavor)) {
                return (String) contents.getTransferData(DataFlavor.stringFlavor);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }
}
