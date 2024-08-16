package org.example;

import org.apache.poi.ss.usermodel.*;

import java.awt.*;

public class OrderProcessor {

    private final RobotHelper robotHelper;

    public OrderProcessor(Robot robot) {
        this.robotHelper = new RobotHelper(robot);
    }

    public void processOrder(Row row) {
        String orderNumber = robotHelper.getCellValueAsString(row.getCell(0));
        String positionNumber = robotHelper.getCellValueAsString(row.getCell(1));
        String deliveryDate = robotHelper.getCellValueAsString(row.getCell(3));
        String confirmationNumber = robotHelper.getCellValueAsString(row.getCell(2));

        System.out.println("Обработка заказа: " + orderNumber);

        // Проверка количества символов в номере заказа
        if (orderNumber.length() == 6) {
            robotHelper.typeText("K");
        } else if (orderNumber.length() == 5) {
            robotHelper.typeText("L");
        } else {
            System.out.println("Ошибка: Номер заказа должен содержать 5 или 6 символов. Остановлена обработка.");
            return;
        }

        robotHelper.pressEnter();

        // Запуск скрипта через F1
        robotHelper.pressF1();

        // Получение значения из буфера обмена
        String clipboardContent = robotHelper.getClipboardContents();

        // Проверка совпадения с ожидаемым значением "311"
        while (!clipboardContent.equals("311")) {
            System.out.println("Значение позиции курсора не равно '311'. Перемещение курсора влево и повторная проверка.");
            robotHelper.pressLeftArrow();

            // Повторная проверка позиции курсора
            robotHelper.pressF1();
            clipboardContent = robotHelper.getClipboardContents();

            if (clipboardContent.equals("2362")) {
                System.out.println("Курсор все еще в положении '2362'. Ввод буквы 'L' и значений '5.0321'.");
                robotHelper.typeText("L");
                robotHelper.pressEnter();
                robotHelper.typeText("5.0321");
                robotHelper.pressEnter();

                System.out.println("Повторная обработка текущего заказа: " + orderNumber);
                processOrder(row);
                return;
            }
        }

        System.out.println("Курсор на правильной позиции. Ввод номера заказа: " + orderNumber);
        robotHelper.typeText(orderNumber);
        robotHelper.pressEnter();

        System.out.println("Ввод номера позиции: " + positionNumber);
        robotHelper.typeText(positionNumber);
        robotHelper.pressEnter();

        System.out.println("Проверка, не доставлен ли заказ.");
        robotHelper.pressF2();

        String clipboardContentWE = robotHelper.getClipboardContents();

        if (clipboardContentWE.contains("Wareneingang")) {
            System.out.println("Заказ уже доставлен. Завершаем обработку.");
            return;
        }

        robotHelper.pressF1();
        clipboardContent = robotHelper.getClipboardContents();

        if (clipboardContent.equals("1374")) {
            System.out.println("Курсор имеет значение '1374'. Нажатие стрелки назад и переход к следующему заказу.");
            robotHelper.pressLeftArrow();
            return;
        }

        System.out.println("Проверка позиции курсора перед вводом даты.");

        robotHelper.pressF1();
        clipboardContent = robotHelper.getClipboardContents();

        while (clipboardContent.equals("2480")) {
            System.out.println("Позиция курсора равна '2480'. Нажатие Enter и повторная проверка.");
            robotHelper.pressEnter();
            robotHelper.pressF1();
            clipboardContent = robotHelper.getClipboardContents();
        }

        if (clipboardContent.equals("936")) {
            System.out.println("Курсор на правильной позиции. Ввод даты поставки: " + deliveryDate);
            robotHelper.typeText(deliveryDate);
            robotHelper.pressEnter();
            robotHelper.pressEnter();
        } else {
            System.out.println("Курсор не на правильной позиции для ввода даты.");
        }

        System.out.println("Ввод номера подтверждения: " + confirmationNumber);
        robotHelper.typeText(confirmationNumber);

        while (!clipboardContent.equals("2375") && !clipboardContent.equals("2376")) {
            robotHelper.pressEnter();
            robotHelper.pressF1();
            clipboardContent = robotHelper.getClipboardContents();

            if (clipboardContent.equals("411")) {
                System.out.println("Курсор имеет значение '411'. Прерывание цикла и переход к следующему заказу.");
                return;
            }
        }

        robotHelper.pressEnter();
        robotHelper.pressF1();
        clipboardContent = robotHelper.getClipboardContents();

        if (clipboardContent.equals("2362")) {
            System.out.println("Курсор имеет значение '2362'. Нажатие стрелки влево.");
            robotHelper.pressLeftArrow();
        }

        System.out.println("Задержка перед обработкой следующего заказа.");
        robotHelper.robot.delay(300);
    }
}
