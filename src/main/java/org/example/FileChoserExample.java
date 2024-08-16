package org.example;

import javafx.application.Platform;
import javafx.embed.swing.JFXPanel;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import javax.swing.*;
import java.io.File;

public class FileChoserExample {

    public static void showFileChooser(JFrame parentFrame, JTextField filePathField) {
        JFXPanel fxPanel = new JFXPanel(); // Инициализация JavaFX среды
        Platform.runLater(() -> {
            FileChooser fileChooser = new FileChooser();
            File file = fileChooser.showOpenDialog(new Stage());
            if (file != null) {
                filePathField.setText(file.getAbsolutePath()); // Устанавливаем путь в текстовое поле
            }
        });
    }
}
