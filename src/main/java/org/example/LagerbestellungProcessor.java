package org.example;

import org.apache.poi.ss.usermodel.Row;

import java.awt.*;

public class LagerbestellungProcessor {

    private final Robot robot;
    private final RobotHelper robotHelper;

    public LagerbestellungProcessor(Robot robot) {
        this.robot = robot;
        this.robotHelper = new RobotHelper(robot);
    }

    public void processLagerComment(Row row) {
        String orderNumber = robotHelper.getCellValueAsString(row.getCell(0));
        String positionNumber = robotHelper.getCellValueAsString(row.getCell(1));

        System.out.println("Обработка комментария для заказа: " + orderNumber);

        // Проверка и установка курсора на позицию 316
        robotHelper.moveToPositionWithLeftArrow("316");

        // Ввод номера заказа
        System.out.println("Ввод номера заказа: " + orderNumber);
        robotHelper.typeText(orderNumber);
        robotHelper.pressEnter();
        robot.delay(300);

        // Проверка и установка курсора на позицию 237
        robotHelper.moveToPositionWithEnter("237");

        // Ввод номера позиции
        System.out.println("Ввод номера позиции: " + positionNumber);
        robotHelper.typeText(positionNumber);
        robotHelper.pressEnter();

        // Проверка позиции 2439 после ввода номера позиции
        robotHelper.checkAndHandlePositionAfterPositionInput();

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
        String comment = "WE-ANFRAGE AN DEN HST GEMAILT";
        System.out.println("Ввод комментария: " + comment);
        robotHelper.typeText(comment);

        // Нажимаем Enter для подтверждения ввода комментария
        robotHelper.pressEnter();

        // Проверка на позицию 2258
        robotHelper.moveToPosition("2258");

        // Ввод команды N на позиции 2258
        System.out.println("Ввод команды N на позиции 2258");
        robotHelper.typeText("N");
        robotHelper.pressEnter();

        robotHelper.checkAndHandlePositionAfterPositionInput();

        // Если не на позиции 2258, возвращаем курсор назад на позицию 316
        robotHelper.moveToPositionWithLeftArrow("316");

        // Задержка перед обработкой следующего комментария
        robot.delay(300);
    }
}
