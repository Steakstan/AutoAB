package org.example;

import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser.ExtensionFilter;

import java.io.File;

public class FileChoserExample extends Application {

    private TextField filePathField;

    @Override
    public void start(Stage primaryStage) {
        // Создаем элементы интерфейса
        filePathField = new TextField();
        filePathField.setPromptText("Выберите файл");

        Button browseButton = new Button("Обзор");
        browseButton.setOnAction(e -> showFileChooser(primaryStage));

        VBox vbox = new VBox(10, filePathField, browseButton);
        vbox.setPadding(new javafx.geometry.Insets(10));

        Scene scene = new Scene(vbox, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.setTitle("File Chooser Example");
        primaryStage.show();
    }

    private void showFileChooser(Stage primaryStage) {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new ExtensionFilter("Excel Files", "*.xlsx"));

        File file = fileChooser.showOpenDialog(primaryStage);
        if (file != null) {
            filePathField.setText(file.getAbsolutePath()); // Устанавливаем путь в текстовое поле
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
