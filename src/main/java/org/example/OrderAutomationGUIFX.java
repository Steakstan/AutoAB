package org.example;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Node;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.shape.Rectangle;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.util.logging.Level;
import java.util.logging.Logger;

public class OrderAutomationGUIFX extends Application {

    private ComboBox<String> operationComboBox;
    private TextField filePathField;
    private OrderAutomation automation;
    private LanguageManager languageManager; // Новый объект для управления языками

    // Переменные для хранения смещения при перетаскивании окна
    private double offsetX, offsetY;

    // Логгер
    private static final Logger LOGGER = Logger.getLogger(OrderAutomationGUIFX.class.getName());

    @Override
    public void start(Stage primaryStage) {
        automation = new OrderAutomation();
        languageManager = new LanguageManager(); // Инициализация LanguageManager

        // Основной контейнер
        BorderPane root = new BorderPane();
        root.getStyleClass().add("root");

        // Панель заголовка
        HBox titleBar = createTitleBar(primaryStage);
        root.setTop(titleBar);

        // Панель для основного контента
        VBox contentPanel = createContentPanel();
        root.setCenter(contentPanel);

        // Создание сцены и установка на Stage
        Scene scene = new Scene(root, 500, 220);
        scene.getStylesheets().add(getClass().getResource("/styles.css").toExternalForm());
        primaryStage.setScene(scene);
        primaryStage.setTitle("Order Automation");
        primaryStage.initStyle(javafx.stage.StageStyle.TRANSPARENT); // Убираем стандартный заголовок
        primaryStage.show();

        // Обработка событий для перетаскивания окна за панель заголовка
        addDragFunctionality(titleBar, primaryStage);
    }

    // Создание панели заголовка
    private HBox createTitleBar(Stage stage) {
        HBox titleBar = new HBox();
        titleBar.getStyleClass().add("title-bar");

        Button minimizeButton = createTitleButton("−", event -> stage.setIconified(true));
        Button closeButton = createTitleButton("×", event -> System.exit(0));

        // Кнопка переключения языка
        Button toggleLanguageButton = createTitleButton("DE/UA", event -> toggleLanguage());
        titleBar.getChildren().addAll(minimizeButton, toggleLanguageButton, closeButton);

        return titleBar;
    }

    // Создание панели основного контента
    private VBox createContentPanel() {
        VBox contentPanel = new VBox(10);
        contentPanel.setPadding(new Insets(10));
        contentPanel.setAlignment(Pos.CENTER_LEFT);
        contentPanel.getStyleClass().add("content-panel");

        Label operationLabel = createLabel(languageManager.getText("operationLabel"));
        operationComboBox = new ComboBox<>();
        operationComboBox.getItems().addAll(
                "AB-Verarbeitung",
                "Kommentare verarbeiten",
                "Kommentare für Lagerbestellung verarbeiten",
                "Liefertermine verarbeiten"
        );
        operationComboBox.getStyleClass().add("combo-box");

        HBox fileSelectionBox = new HBox(10);
        fileSelectionBox.setAlignment(Pos.CENTER_LEFT);

        Label filePathLabel = createLabel(languageManager.getText("filePathLabel"));
        filePathField = new TextField();
        filePathField.getStyleClass().add("text-field");
        filePathField.setPrefWidth(300);

        Button browseButton = createRoundedButton(languageManager.getText("browseButton"));
        browseButton.setOnAction(e -> showFileChooser());

        fileSelectionBox.getChildren().addAll(filePathField, browseButton);

        Button startButton = createRoundedButton(languageManager.getText("startButton"));
        startButton.setOnAction(e -> startProcessing());

        contentPanel.getChildren().addAll(operationLabel, operationComboBox, filePathLabel, fileSelectionBox, startButton);

        return contentPanel;
    }

    // Показ диалога выбора файла
    private void showFileChooser() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
        fileChooser.setTitle(languageManager.getText("fileChooserTitle"));
        java.io.File selectedFile = fileChooser.showOpenDialog(null);
        if (selectedFile != null) {
            filePathField.setText(selectedFile.getAbsolutePath());
        }
    }

    // Запуск обработки
    private void startProcessing() {
        int choice = operationComboBox.getSelectionModel().getSelectedIndex() + 1;
        String filePath = filePathField.getText();

        if (filePath.isEmpty()) {
            showAlert(Alert.AlertType.ERROR, languageManager.getText("errorTitle"), languageManager.getText("errorMessage"));
            return;
        }

        try {
            automation.processFile(choice, filePath);
            showAlert(Alert.AlertType.INFORMATION, languageManager.getText("successTitle"), languageManager.getText("successMessage"));
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "Error during file processing", e);
            showAlert(Alert.AlertType.ERROR, languageManager.getText("errorTitle"), languageManager.getText("processingErrorMessage"));
        }
    }

    // Создание кнопки с округлыми углами
    private Button createRoundedButton(String text) {
        Button button = new Button(text);
        button.getStyleClass().add("rounded-button");
        button.setShape(new Rectangle(20, 20));
        return button;
    }

    // Создание кнопок заголовка (закрытие и сворачивание)
    private Button createTitleButton(String text, javafx.event.EventHandler<javafx.event.ActionEvent> handler) {
        Button button = new Button(text);
        button.getStyleClass().add("title-button");
        button.setOnAction(handler);
        return button;
    }

    // Создание метки
    private Label createLabel(String text) {
        Label label = new Label(text);
        label.getStyleClass().add("label-white");
        return label;
    }

    // Показ предупреждений
    private void showAlert(Alert.AlertType alertType, String title, String message) {
        Alert alert = new Alert(alertType);
        alert.setTitle(title);
        alert.setContentText(message);
        alert.showAndWait();
    }

    // Добавление перетаскивания окна
    private void addDragFunctionality(HBox titleBar, Stage primaryStage) {
        titleBar.setOnDragDetected(event -> {
            offsetX = event.getSceneX();
            offsetY = event.getSceneY();
            titleBar.startFullDrag();
        });

        titleBar.setOnMouseDragged(event -> {
            primaryStage.setX(event.getScreenX() - offsetX);
            primaryStage.setY(event.getScreenY() - offsetY);
        });
    }

    // Переключение языка интерфейса
    private void toggleLanguage() {
        languageManager.toggleLanguage();

        // Получаем панель контента
        VBox contentPanel = (VBox) ((BorderPane) operationComboBox.getScene().getRoot()).getCenter();

        if (contentPanel != null) {
            // Переводим метку "Wählen Sie den Verarbeitungstyp:"
            Node node1 = contentPanel.getChildren().get(0);
            if (node1 instanceof Label operationLabel) {
                operationLabel.setText(languageManager.getText("operationLabel"));
            }

            // Переводим элементы ComboBox
            operationComboBox.getItems().setAll(
                    languageManager.isGerman()
                            ? new String[]{"AB-Verarbeitung", "Kommentare verarbeiten", "Kommentare für Lagerbestellung verarbeiten", "Liefertermine verarbeiten"}
                            : new String[]{"Обробка AB", "Обробка коментарів", "Обробка коментарів для замовлення на склад", "Обробка термінів поставки"}
            );

            // Переводим метку "Wählen Sie die Excel-Datei:"
            Node node2 = contentPanel.getChildren().get(2);
            if (node2 instanceof Label filePathLabel) {
                filePathLabel.setText(languageManager.getText("filePathLabel"));
            }

            // Переводим кнопку "Verarbeitung starten"
            Node node3 = contentPanel.getChildren().get(4);
            if (node3 instanceof Button startButton) {
                startButton.setText(languageManager.getText("startButton"));
            }

            // Переводим кнопку "Durchsuchen"
            Node node4 = ((HBox) contentPanel.getChildren().get(3)).getChildren().get(1);
            if (node4 instanceof Button browseButton) {
                browseButton.setText(languageManager.getText("browseButton"));
            }
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
