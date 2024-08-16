package org.example;

import org.apache.poi.ss.usermodel.Row;

import java.awt.*;

public class DeliveryDateProcessor {

    private final RobotHelper robotHelper;

    public DeliveryDateProcessor(Robot robot) {
        this.robotHelper = new RobotHelper(robot);
    }

    public void processDeliveryDate(Row row) {
        String orderNumber = robotHelper.getCellValueAsString(row.getCell(0));
        String positionNumber = robotHelper.getCellValueAsString(row.getCell(1));
        String deliveryDate = robotHelper.getCellValueAsString(row.getCell(2));

        startProcessing:

        while (true) {
            System.out.println("Обработка даты поставки для заказа: " + orderNumber);

            if (orderNumber.length() == 6) {
                System.out.println("Номер заказа содержит 6 символов. Ввод буквы 'K'.");
                robotHelper.typeText("K");
            } else if (orderNumber.length() == 5) {
                System.out.println("Номер заказа содержит 5 символов. Ввод буквы 'L'.");
                robotHelper.typeText("L");
            } else {
                System.out.println("Ошибка: Номер заказа должен содержать 5 или 6 символов. Остановлена обработка.");
                return;
            }

            robotHelper.pressEnter();
            robotHelper.pressF1();
            String clipboardContent = robotHelper.getClipboardContents();

            while (!clipboardContent.equals("311")) {
                robotHelper.pressLeftArrow();
                robotHelper.pressF1();
                clipboardContent = robotHelper.getClipboardContents();

                if (clipboardContent.equals("2362")) {
                    robotHelper.typeText("L");
                    robotHelper.pressEnter();
                    robotHelper.typeText("5.0321");
                    robotHelper.pressEnter();
                    continue startProcessing;
                }
            }

            robotHelper.typeText(orderNumber);
            robotHelper.pressEnter();
            robotHelper.typeText(positionNumber);
            robotHelper.pressEnter();

            robotHelper.pressF2();
            clipboardContent = robotHelper.getClipboardContents();

            if (clipboardContent.contains("Wareneingang")) {
                System.out.println("Заказ уже доставлен. Прекращение обработки.");
                return;
            }

            robotHelper.pressF1();
            clipboardContent = robotHelper.getClipboardContents();

            if (clipboardContent.equals("1374")) {
                robotHelper.pressLeftArrow();
                return;
            }

            robotHelper.pressF1();
            clipboardContent = robotHelper.getClipboardContents();

            while (clipboardContent.equals("2480")) {
                robotHelper.pressEnter();
                robotHelper.pressF1();
                clipboardContent = robotHelper.getClipboardContents();
            }

            if (clipboardContent.equals("936")) {
                robotHelper.typeText(deliveryDate);
                robotHelper.pressEnter();
                robotHelper.pressEnter();
            } else {
                System.out.println("Курсор не на правильной позиции для ввода даты.");
            }

            System.out.println("Дата поставки успешно обновлена для заказа: " + orderNumber);
            break;
        }
    }
}
