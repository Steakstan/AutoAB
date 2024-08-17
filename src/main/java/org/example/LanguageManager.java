package org.example;

import java.util.HashMap;
import java.util.Map;

public class LanguageManager {

    private boolean isGerman = true; // Флаг для отслеживания текущего языка
    private final Map<String, String> germanTexts = new HashMap<>();
    private final Map<String, String> ukrainianTexts = new HashMap<>();

    public LanguageManager() {
        // Инициализация текстов на немецком языке
        germanTexts.put("operationLabel", "Wählen Sie den Verarbeitungstyp:");
        germanTexts.put("filePathLabel", "Wählen Sie die Excel-Datei:");
        germanTexts.put("startButton", "Verarbeitung starten");
        germanTexts.put("browseButton", "Durchsuchen");
        germanTexts.put("fileChooserTitle", "Wählen Sie eine Excel-Datei");
        germanTexts.put("errorTitle", "Fehler");
        germanTexts.put("errorMessage", "Bitte wählen Sie eine Excel-Datei aus.");
        germanTexts.put("successTitle", "Erfolg");
        germanTexts.put("successMessage", "Die Verarbeitung wurde erfolgreich abgeschlossen.");
        germanTexts.put("processingErrorMessage", "Fehler bei der Verarbeitung der Datei.");

        // Инициализация текстов на украинском языке
        ukrainianTexts.put("operationLabel", "Виберіть тип обробки:");
        ukrainianTexts.put("filePathLabel", "Виберіть файл Excel:");
        ukrainianTexts.put("startButton", "Запустити обробку");
        ukrainianTexts.put("browseButton", "Огляд");
        ukrainianTexts.put("fileChooserTitle", "Виберіть файл Excel");
        ukrainianTexts.put("errorTitle", "Помилка");
        ukrainianTexts.put("errorMessage", "Будь ласка, виберіть файл Excel.");
        ukrainianTexts.put("successTitle", "Успіх");
        ukrainianTexts.put("successMessage", "Обробка успішно завершена.");
        ukrainianTexts.put("processingErrorMessage", "Помилка при обробці файлу.");
    }

    // Метод для переключения языка
    public void toggleLanguage() {
        isGerman = !isGerman;
    }

    // Метод для получения текста на текущем языке
    public String getText(String key) {
        return isGerman ? germanTexts.get(key) : ukrainianTexts.get(key);
    }

    public boolean isGerman() {
        return isGerman;
    }
}
