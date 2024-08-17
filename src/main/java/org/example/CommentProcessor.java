package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.KeyEvent;

public class CommentProcessor {

    private final Robot robot;

    public CommentProcessor(Robot robot) {
        this.robot = robot;
    }

    public void processComment(Row row) {
        String orderNumber = getCellValueAsString(row.getCell(0));
        String positionNumber = getCellValueAsString(row.getCell(1));

        System.out.println("Обработка комментария для заказа: " + orderNumber);

        // Проверка и установка курсора на позицию 313
        moveToPositionWithLeftArrow();

        // Ввод номера заказа
        System.out.println("Ввод номера заказа: " + orderNumber);
        typeText(orderNumber);
        pressEnter();
        robot.delay(300);

        // Проверка и установка курсора на позицию 2310
        moveToPositionWithEnter("2310");

        // Ввод номера позиции
        System.out.println("Ввод номера позиции: " + positionNumber);
        typeText(positionNumber);
        pressEnter();

        // Проверка позиции 2439 после ввода номера позиции
        checkAndHandlePositionAfterPositionInput();

        // Проверка и установка курсора на позицию 237
        moveToPosition("237");

        // Ввод команды Z
        System.out.println("Ввод команды Z");
        typeText("Z");
        pressEnter();

        // Переход на позицию 222
        moveToPositionWithEnter("222");

        // Проверка и установка курсора на позицию 2210
        moveToPosition("2210");

        // Ввод комментария
        String comment = "*WE-ANFRAGE AN DEN HST GEMAILT";
        System.out.println("Ввод комментария: " + comment);
        typeText(comment);

        // Нажимаем Enter для подтверждения ввода комментария
        pressEnter();

        // Если не на позиции 2210, возвращаем курсор назад на позицию 313
        moveToPositionWithLeftArrow();

        // Задержка перед обработкой следующего комментария
        robot.delay(300);
    }

    private void checkAndHandlePositionAfterPositionInput() {
        String cursorPosition = getCursorPosition();

        while ("2440".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2440. Нажатие Enter.");
            pressEnter();
            cursorPosition = getCursorPosition(); // Обновляем позицию курсора после нажатия Enter
        }

        if (cursorPosition.equals("2439")) {
            System.out.println("Курсор на позиции 2439. Нажатие N и Enter.");
            typeText("N");
            pressEnter();
        } else {
            System.out.println("Курсор не на позиции 2439 или 2440. Текущая позиция: " + cursorPosition);
        }
    }

    private String getCursorPosition() {
        pressF1();
        return getClipboardContents();
    }

    private void moveToPositionWithEnter(String expectedPosition) {
        while (!moveToPosition(expectedPosition)) {
            pressEnter();
        }
        pressEnter();
    }

    private void moveToPositionWithLeftArrow() {
        while (!moveToPosition("313")) {
            pressLeftArrow();
        }
    }

    private boolean moveToPosition(String expectedPosition) {
        pressF1();
        String clipboardContent = getClipboardContents();

        if (clipboardContent.equals(expectedPosition)) {
            System.out.println("Курсор установлен на позицию " + expectedPosition + ".");
            return true;
        } else {
            System.out.println("Курсор не на позиции " + expectedPosition + ".");
            return false;
        }
    }

    private void pressRightArrow() {
        System.out.println("Нажатие стрелки вправо");
        robot.keyPress(KeyEvent.VK_RIGHT);
        robot.keyRelease(KeyEvent.VK_RIGHT);
        robot.delay(300); // Задержка после нажатия стрелки вправо
    }

    private String getCellValueAsString(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }

    private void typeText(String text) {
        for (char c : text.toCharArray()) {
            try {
                if (Character.isUpperCase(c)) {
                    robot.keyPress(KeyEvent.VK_SHIFT);
                }

                if (Character.isLetterOrDigit(c)) {
                    int keyCode = KeyEvent.getExtendedKeyCodeForChar(c);
                    robot.keyPress(keyCode);
                    robot.keyRelease(keyCode);
                } else {
                    // Обработка специальных символов
                    switch (c) {
                        case ' ':
                            robot.keyPress(KeyEvent.VK_SPACE);
                            robot.keyRelease(KeyEvent.VK_SPACE);
                            break;
                        case '-':
                            robot.keyPress(KeyEvent.VK_MINUS);
                            robot.keyRelease(KeyEvent.VK_MINUS);
                            break;
                        case '*':
                            robot.keyPress(KeyEvent.VK_SHIFT);
                            robot.keyPress(KeyEvent.VK_8); // * это Shift + 8
                            robot.keyRelease(KeyEvent.VK_8);
                            robot.keyRelease(KeyEvent.VK_SHIFT);
                            break;
                        case '.':
                            robot.keyPress(KeyEvent.VK_PERIOD);
                            robot.keyRelease(KeyEvent.VK_PERIOD);
                            break;
                        default:
                            System.out.println("Неподдерживаемый символ: " + c);
                            break;
                    }
                }

                if (Character.isUpperCase(c)) {
                    robot.keyRelease(KeyEvent.VK_SHIFT);
                }

            } catch (IllegalArgumentException e) {
                System.out.println("Ошибка ввода символа: " + c + ". " + e.getMessage());
            }
        }
    }


    private void pressEnter() {
        System.out.println("Нажатие Enter");
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
        robot.delay(300); // Задержка после нажатия Enter
    }

    private void pressF1() {
        System.out.println("Нажатие F1");
        robot.keyPress(KeyEvent.VK_F1);
        robot.keyRelease(KeyEvent.VK_F1);
        robot.delay(300); // Задержка после нажатия F1
    }

    private void pressLeftArrow() {
        System.out.println("Нажатие стрелки влево");
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