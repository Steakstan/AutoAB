package org.example;

import org.apache.poi.ss.usermodel.Row;

import java.awt.*;

public class CommentProcessor {

    private final RobotHelper robotHelper;

    public CommentProcessor(Robot robot) {
        this.robotHelper = new RobotHelper(robot);
    }

    public void processComment(Row row) {
        String orderNumber = robotHelper.getCellValueAsString(row.getCell(0));
        String positionNumber = robotHelper.getCellValueAsString(row.getCell(1));

        System.out.println("Обработка комментария для заказа: " + orderNumber);

        // Проверка и установка курсора на позицию 313
        robotHelper.moveToPositionWithLeftArrow("313");

        // Ввод номера заказа
        System.out.println("Ввод номера заказа: " + orderNumber);
        robotHelper.typeText(orderNumber);
        robotHelper.pressEnter();
        robotHelper.robot.delay(300);

        // Проверка и установка курсора на позицию 2310
        robotHelper.moveToPositionWithEnter("2310");

        // Ввод номера позиции
        System.out.println("Ввод номера позиции: " + positionNumber);
        robotHelper.typeText(positionNumber);
        robotHelper.pressEnter();

        // Проверка позиции 2439 после ввода номера позиции
        checkAndHandlePositionAfterPositionInput();

        // Проверка и установка курсора на позицию 237
        robotHelper.moveToPosition("237");

        // Ввод команды Z
        System.out.println("Ввод команды Z");
        robotHelper.typeText("Z");
        robotHelper.pressEnter();

        // Переход на позицию 222
        robotHelper.moveToPositionWithEnter("222");

        // Проверка и установка курсора на позицию 2210
        robotHelper.moveToPosition("2210");

        // Ввод комментария
        String comment = "*WE-ANFRAGE AN DEN HST GEMAILT";
        System.out.println("Ввод комментария: " + comment);
        robotHelper.typeText(comment);

        // Нажимаем Enter для подтверждения ввода комментария
        robotHelper.pressEnter();

        // Если не на позиции 2210, возвращаем курсор назад на позицию 313
        robotHelper.moveToPositionWithLeftArrow("313");

        // Задержка перед обработкой следующего комментария
        robotHelper.robot.delay(300);
    }

    private void checkAndHandlePositionAfterPositionInput() {
        String cursorPosition = robotHelper.getCursorPosition();

        while ("2440".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2440. Нажатие Enter.");
            robotHelper.pressEnter();
            cursorPosition = robotHelper.getCursorPosition(); // Обновляем позицию курсора после нажатия Enter
        }

        if (cursorPosition.equals("2439")) {
            System.out.println("Курсор на позиции 2439. Нажатие N и Enter.");
            robotHelper.typeText("N");
            robotHelper.pressEnter();
        } else {
            System.out.println("Курсор не на позиции 2439 или 2440. Текущая позиция: " + cursorPosition);
        }
    }
}
