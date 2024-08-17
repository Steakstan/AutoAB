package org.example;

import org.apache.poi.ss.usermodel.Row;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.event.KeyEvent;
import org.apache.poi.ss.usermodel.*;

public class DeliveryDateProcessor {

    private final Robot robot;

    public DeliveryDateProcessor(Robot robot) {
        this.robot = robot;
    }

    public void processDeliveryDate(Row row) {
        String orderNumber = getCellValueAsString(row.getCell(0));
        String positionNumber = getCellValueAsString(row.getCell(1));
        String deliveryDate = getCellValueAsString(row.getCell(2));

        // Метка для возврата к началу обработки заказа
        startProcessing:

        while (true) {
            System.out.println("Обработка даты поставки для заказа: " + orderNumber);

            // Проверка номера заказа
            if (orderNumber.length() == 6) {
                System.out.println("Номер заказа содержит 6 символов. Ввод буквы 'K'.");
                typeText("K");
            } else if (orderNumber.length() == 5) {
                System.out.println("Номер заказа содержит 5 символов. Ввод буквы 'L'.");
                typeText("L");
            } else {
                System.out.println("Ошибка: Номер заказа должен содержать 5 или 6 символов. Остановлена обработка.");
                return;
            }

            pressEnter();
            System.out.println("Нажата клавиша Enter.");

            // Проверка позиции курсора через F1
            System.out.println("Проверка позиции курсора через F1.");
            pressF1();
            String clipboardContent = getClipboardContents();
            while (!clipboardContent.equals("311")) {
                System.out.println("Значение позиции курсора не равно '311'. Перемещение курсора влево и повторная проверка.");
                pressLeftArrow();

                // Повторная проверка позиции курсора
                pressF1();
                clipboardContent = getClipboardContents();

                if (clipboardContent.equals("2362")) {
                    System.out.println("Курсор в положении '2362'. Ввод буквы 'L' и значений '5.0321'.");
                    typeText("L");
                    pressEnter();
                    typeText("5.0321");
                    pressEnter();

                    System.out.println("Повторная обработка текущего заказа: " + getCellValueAsString(row.getCell(0)));
                    processDeliveryDate(row); // Повторный вызов метода для обработки текущего заказа
                    return; // Завершение текущей обработки
                }
            }
            System.out.println("Курсор установлен на позицию 311.");

            // Ввод номера заказа и позиции
            System.out.println("Ввод номера заказа: " + orderNumber);
            typeText(orderNumber);
            pressEnter();

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
            // Проверка позиции курсора после ввода позиции
            System.out.println("Проверка позиции курсора через F1 после ввода позиции.");
            pressF1();
            clipboardContent = getClipboardContents();
            if (!clipboardContent.equals("1374")) {
                System.out.println("Курсор не на позиции 1374. Переход к следующему заказу.");
                return; // Прерываем обработку текущего заказа
            }

// Если курсор на позиции 1374
            System.out.println("Курсор установлен на позицию 1374. Ввод буквы 'N' и нажатие Enter.");
            typeText("N");
            pressEnter();
            pressF1();
            clipboardContent = getClipboardContents();

// Продолжаем нажатие Enter до позиции 936
            System.out.println("Продолжаем нажатие Enter до достижения позиции 936.");
            // Для хранения предыдущей позиции курсора
            int consecutive411Count = 0; // Счетчик для отслеживания повторных появлений позиции 411

            while (!clipboardContent.equals("936")) {
                pressEnter();
                pressF1();
                clipboardContent = getClipboardContents();
                System.out.println("Нажимаем энтер. Номер позиции " + clipboardContent);

                // Проверка на повторную позицию 411
                if (clipboardContent.equals("411")) {
                    consecutive411Count++;
                    System.out.println("Курсор на позиции 411. Счетчик повторных появлений: " + consecutive411Count);
                    if (consecutive411Count == 2) {
                        System.out.println("Курсор дважды подряд на позиции 411. Сброс обработки заказа.");
                        return; // Возвращаемся к началу обработки того же заказа
                    }
                } else {
                    consecutive411Count = 0; // Сбрасываем счетчик если позиция отличается от 411
                }
            }
            System.out.println("Курсор установлен на позицию 936.");

// Ввод даты поставки
            System.out.println("Ввод даты поставки: " + deliveryDate);
            typeText(deliveryDate);
            pressEnter();
            pressF1();
            clipboardContent = getClipboardContents();


            // Нажимаем "T" и Enter до позиции 2375
            System.out.println("Ввод буквы 'T' и нажатие Enter до достижения позиции 2375.");
            typeText("T");
            pressEnter();
            if(orderNumber.length()==5) {
                robot.delay(5000);
            }
            label:
            while (true) {
                pressEnter();
                pressF1();
                clipboardContent = getClipboardContents();
                System.out.println("нажатие Enter " + clipboardContent);

                switch (clipboardContent) {
                    case "2375":
                        System.out.println("Курсор установлен на позицию 2375.");
                        break label;
                    case "2376":
                        System.out.println("Курсор установлен на позицию 2376.");
                        break label;
                    case "411":
                        System.out.println("Курсор установлен на позицию 411. Прерывание цикла и повторная обработка заказа.");
                        continue startProcessing;
                }

            }

            // Вводим "Z" и нажимаем Enter до позиции 2212
            System.out.println("Ввод буквы 'Z' и нажатие Enter до достижения позиции 2212.");
            typeText("Z");
            pressEnter();
            pressF1();
            clipboardContent = getClipboardContents();
            while (!clipboardContent.equals("2212")&&!clipboardContent.equals("2211")) {
                pressEnter();
                pressF1();
                clipboardContent = getClipboardContents();
            }
            System.out.println("Курсор установлен на позицию 2212.");

            // Извлекаем первые два числа из даты поставки для комментария
            String kwWeek = extractWeekFromDeliveryDate(deliveryDate);

            // Вставляем комментарий с использованием даты поставки
            String comment = "DEM HST NACH WIRD DIE WARE IN KW " + kwWeek + " ZUGESTELLT";
            System.out.println("Ввод комментария: " + comment);
            typeText(comment);
            pressEnter();

            // Проверка позиции 2274
            System.out.println("Проверка позиции курсора через F1.");
            pressF1();
            clipboardContent = getClipboardContents();
            if (clipboardContent.equals("2274")) {
                System.out.println("Курсор установлен на позицию 2274. Нажатие Enter.");
                pressEnter();
            }
            if (clipboardContent.equals("2273")) {
                System.out.println("Курсор установлен на позицию 2274. Нажатие Enter.");
                pressEnter();
            }


            // Проверка позиции 222
            System.out.println("Проверка позиции 222.");
            pressF1();
            clipboardContent = getClipboardContents();
            if (clipboardContent.equals("222")) {
                System.out.println("Курсор установлен на позицию 222.");
                // Возвращаемся к позиции 2375
                System.out.println("Возвращение к позиции 2375 путем нажатия стрелки влево.");
                while (!clipboardContent.equals("2375")&&!clipboardContent.equals("2376")) {
                    pressLeftArrow();
                    pressF1();
                    clipboardContent = getClipboardContents();
                }
                pressEnter();
            }

            // Проверка на позицию 2362 для завершения
            System.out.println("Проверка позиции 2362.");
            pressF1();
            clipboardContent = getClipboardContents();
            if (clipboardContent.equals("2362")) {
                System.out.println("Курсор установлен на позицию 2362. Нажатие стрелки влево.");
                pressLeftArrow();
            }

            System.out.println("Задержка перед обработкой следующего заказа.");
            robot.delay(300); // Задержка между итерациями

            break; // Выходим из цикла после успешной обработки заказа
        }
    }


    private String getCellValueAsString(Cell cell) {
        // Логика аналогична OrderProcessor
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

    private String extractWeekFromDeliveryDate(String deliveryDate) {
        // Предполагается, что deliveryDate имеет формат "dd.MM.yyyy" или подобный
        if (deliveryDate != null && deliveryDate.length() >= 2) {
            return deliveryDate.substring(0, 2);  // Извлекаем первые два символа
        }
        return "??";  // Возвращаем значение по умолчанию, если дата некорректна
    }

    private void typeText(String text) {
        // Логика аналогична OrderProcessor
        for (char c : text.toCharArray()) {
            if (Character.isUpperCase(c)) {
                robot.keyPress(KeyEvent.VK_SHIFT);
            }
            int keyCode = KeyEvent.getExtendedKeyCodeForChar(Character.toUpperCase(c));
            if (KeyEvent.CHAR_UNDEFINED == keyCode) {
                throw new RuntimeException("Key code not found for character '" + c + "'");
            }
            robot.keyPress(keyCode);
            robot.keyRelease(keyCode);

            if (Character.isUpperCase(c)) {
                robot.keyRelease(KeyEvent.VK_SHIFT);
            }
        }
    }

    private void pressEnter() {
        robot.keyPress(KeyEvent.VK_ENTER);
        robot.keyRelease(KeyEvent.VK_ENTER);
        robot.delay(300);
    }

    private void pressF1() {
        robot.keyPress(KeyEvent.VK_F1);
        robot.keyRelease(KeyEvent.VK_F1);
        robot.delay(300);
    }

    private void pressF2() {
        robot.keyPress(KeyEvent.VK_F2);
        robot.keyRelease(KeyEvent.VK_F2);
        robot.delay(300);
    }

    private void pressLeftArrow() {
        robot.keyPress(KeyEvent.VK_LEFT);
        robot.keyRelease(KeyEvent.VK_LEFT);
        robot.delay(300);
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