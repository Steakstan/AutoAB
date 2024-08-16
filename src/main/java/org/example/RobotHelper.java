package org.example;

import org.apache.poi.ss.usermodel.Cell;
import java.awt.*;
import java.awt.datatransfer.*;
import java.awt.event.KeyEvent;

public class RobotHelper {

    final Robot robot;

    public RobotHelper(Robot robot) {
        this.robot = robot;
    }

    public void pressEnter() {
        System.out.println("Нажатие Enter");
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
        robot.delay(300);
    }

    public void pressF1() {
        System.out.println("Нажатие F1");
        robot.keyPress(KeyEvent.VK_F1);
        robot.keyRelease(KeyEvent.VK_F1);
        robot.delay(300);
    }

    public void pressF2() {
        robot.keyPress(KeyEvent.VK_F2);
        robot.keyRelease(KeyEvent.VK_F2);
        robot.delay(300);
    }

    public void pressLeftArrow() {
        System.out.println("Нажатие стрелки влево");
        robot.keyPress(KeyEvent.VK_LEFT);
        robot.keyRelease(KeyEvent.VK_LEFT);
        robot.delay(300);
    }

    public void pressRightArrow() {
        System.out.println("Нажатие стрелки вправо");
        robot.keyPress(KeyEvent.VK_RIGHT);
        robot.keyRelease(KeyEvent.VK_RIGHT);
        robot.delay(300);
    }

    public String getClipboardContents() {
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

    public String getCellValueAsString(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((int) cell.getNumericCellValue());
            default -> "";
        };
    }

    public void typeText(String text) {
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

    public String getCursorPosition() {
        pressF1();
        return getClipboardContents();
    }

    public void checkAndHandlePositionAfterPositionInput() {
        String cursorPosition = getCursorPosition();

        while ("2440".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2440. Нажатие Enter.");
            pressEnter();
            cursorPosition = getCursorPosition(); // Обновляем позицию курсора после нажатия Enter
        }
        while ("2480".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2480. Нажатие Enter.");
            pressEnter();
            cursorPosition = getCursorPosition(); // Обновляем позицию курсора после нажатия Enter
        }

        if ("2439".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2439. Нажатие N и Enter.");
            typeText("N");
            pressEnter();
            cursorPosition = getCursorPosition(); // Обновляем позицию курсора после нажатия Enter
        }

        // Проверка на позицию 2259 после ввода N
        if ("2259".equals(cursorPosition)) {
            System.out.println("Курсор на позиции 2259. Повторное нажатие N и Enter.");
            typeText("N");
            pressEnter();
        } else {
            System.out.println("Курсор не на позиции 2439 или 2440. Текущая позиция: " + cursorPosition);
        }
    }

    public boolean moveToPosition(String expectedPosition) {
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

    public void moveToPositionWithEnter(String expectedPosition) {
        while (!moveToPosition(expectedPosition)) {
            pressEnter();
        }
        pressEnter();
    }

    public void moveToPositionWithLeftArrow(String expectedPosition) {
        while (!moveToPosition(expectedPosition)) {
            pressLeftArrow();
        }
    }
}
