package org.example;

import org.apache.poi.ss.usermodel.*;

import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.KeyEvent;

public class OrderProcessor {

    private final Robot robot;

    public OrderProcessor(Robot robot) {
        this.robot = robot;
    }

    public void processOrder(Row row) {
        String orderNumber = getCellValueAsString(row.getCell(0));
        String positionNumber = getCellValueAsString(row.getCell(1));
        String deliveryDate = getCellValueAsString(row.getCell(3));
        String confirmationNumber = getCellValueAsString(row.getCell(2));

        System.out.println("Обработка заказа: " + orderNumber);

        // Проверка количества символов в номере заказа
        if (orderNumber.length() == 6) {
            typeText("K");
        } else if (orderNumber.length() == 5) {
            typeText("L");
        } else {
            System.out.println("Ошибка: Номер заказа должен содержать 5 или 6 символов. Остановлена обработка.");
            return; // Остановить обработку, если номер заказа не правильный
        }

        pressEnter();

        // Запуск скрипта через F1
        pressF1();

        // Получение значения из буфера обмена
        String clipboardContent = getClipboardContents();

        // Проверка совпадения с ожидаемым значением "311"
        while (!clipboardContent.equals("311")) {
            System.out.println("Значение позиции курсора не равно '311'. Перемещение курсора влево и повторная проверка.");
            pressLeftArrow();

            // Повторная проверка позиции курсора
            pressF1();
            clipboardContent = getClipboardContents();

            if (clipboardContent.equals("2362")) {
                System.out.println("Курсор все еще в положении '2362'. Ввод буквы 'L' и значений '5.0321'.");
                typeText("L");
                pressEnter();
                typeText("5.0321");
                pressEnter();

                System.out.println("Повторная обработка текущего заказа: " + getCellValueAsString(row.getCell(0)));
                processOrder(row); // Повторный вызов метода для обработки текущего заказа
                return; // Завершение текущей обработки
            }
        }

        System.out.println("Курсор на правильной позиции. Ввод номера заказа: " + orderNumber);
        typeText(orderNumber);

        pressEnter();
        robot.delay(300);

        System.out.println("Ввод номера позиции: " + positionNumber);
        typeText(positionNumber);
        pressEnter();

        System.out.println("Проверка, не доставлен ли заказ.");
        pressF2(); // Нажимаем F2 для запуска скрипта, который копирует строку

        String clipboardContentWE = getClipboardContents(); // Получаем содержимое буфера обмена

        if (clipboardContentWE.contains("Wareneingang")) {
            System.out.println("Заказ уже доставлен. Завершаем обработку.");
            return; // Завершаем обработку текущего заказа и переходим к следующему
        }

        // Проверка позиции курсора после ввода номера позиции
        pressF1();
        clipboardContent = getClipboardContents();

        if (clipboardContent.equals("1374")) {
            System.out.println("Курсор имеет значение '1374'. Нажатие стрелки назад и переход к следующему заказу.");
            pressLeftArrow(); // Возвращаемся назад
            return; // Завершение обработки текущего заказа
        }

        System.out.println("Проверка позиции курсора перед вводом даты.");

        // Запуск скрипта через F1 для проверки позиции перед вводом даты
        pressF1();

        // Получение значения из буфера обмена
        clipboardContent = getClipboardContents();

        // Проверка позиции курсора
        while (clipboardContent.equals("2480")) {
            System.out.println("Позиция курсора равна '2480'. Нажатие Enter и повторная проверка.");
            pressEnter();
            pressF1();
            clipboardContent = getClipboardContents();
        }

        if (clipboardContent.equals("936")) {
            System.out.println("Курсор на правильной позиции. Ввод даты поставки: " + deliveryDate);
            typeText(deliveryDate);
            pressEnter();
            pressEnter();
        } else {
            System.out.println("Курсор не на правильной позиции для ввода даты.");
        }

        System.out.println("Ввод номера подтверждения: " + confirmationNumber);
        typeText(confirmationNumber);

        // Проверка позиции курсора для подтверждения
        while (!clipboardContent.equals("2375")&&!clipboardContent.equals("2376")) {
            pressEnter();
            pressF1();
            clipboardContent = getClipboardContents();

            // Если значение курсора равно "411", возвращаемся к шагу ввода номера заказа
            if (clipboardContent.equals("411")) {
                System.out.println("Курсор имеет значение '411'. Прерывание цикла и переход к следующему заказу.");
                return; // Завершение обработки текущего заказа
            }
        }

        // Проверка после нажатия Enter, если позиция курсора стала равна 2362
        pressEnter();
        pressF1();
        clipboardContent = getClipboardContents();

        if (clipboardContent.equals("2362")) {
            System.out.println("Курсор имеет значение '2362'. Нажатие стрелки влево.");
            pressLeftArrow();
        }


        System.out.println("Задержка перед обработкой следующего заказа.");
        robot.delay(300); // Задержка между итерациями
    }

    private String getCellValueAsString(Cell cell) {
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

    private void typeText(String text) {
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

    private void pressEnter() {
        System.out.println("Нажатие Enter.");
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
        robot.delay(300); // Добавляем задержку в 400 миллисекунд после каждого нажатия Enter
    }

    private void pressF1() {
        robot.keyPress(KeyEvent.VK_F1);
        robot.keyRelease(KeyEvent.VK_F1);
        robot.delay(300); // Задержка, чтобы убедиться, что команда выполнена
    }

    private void pressF2() {
        robot.keyPress(KeyEvent.VK_F2);
        robot.keyRelease(KeyEvent.VK_F2);
        robot.delay(300);
    }

    private void pressLeftArrow() {
        System.out.println("Нажатие стрелки влево.");
        robot.keyPress(KeyEvent.VK_LEFT);
        robot.keyRelease(KeyEvent.VK_LEFT);
        robot.delay(300); // Задержка после нажатия стрелки влево
    }

    private String getClipboardContents() {
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