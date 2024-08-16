package org.example;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.awt.event.*;
import java.awt.geom.RoundRectangle2D;

public class OrderAutomationGUI extends JFrame {


    private JComboBox<String> operationComboBox;
    private JTextField filePathField;
    private JButton browseButton;
    private JButton startButton;
    private OrderAutomation automation;
    private JPanel titleBar;
    private JButton closeButton;
    private JButton minimizeButton;

    public OrderAutomationGUI() {
        super("Order Automation");

        this.automation = new OrderAutomation();

        // Создаем объект шрифта Segoe UI с размером 11
        Font segoeFont = new Font("Segoe UI", Font.PLAIN, 13);

        // Цвета
        Color backgroundColorFon = new Color(33, 33, 33);
        Color backgroundColor = new Color(11, 11, 11);
        Color backgroundButtonColor = new Color(40, 40, 40);
        Color buttonHoverColor = new Color(55, 55, 55);
        Color textColor = Color.WHITE;
        Color inputColor = backgroundColorFon;

        // Установка макета и стилей
        setUndecorated(true);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(500, 180);
        setResizable(false);
        setLocationRelativeTo(null);

        // Скругление углов окна
        setShape(new RoundRectangle2D.Double(0, 0, getWidth(), getHeight(), 20, 20));

        // Панель заголовка
        titleBar = new JPanel();
        titleBar.setLayout(new BorderLayout());
        titleBar.setBackground(backgroundColor);
        titleBar.setPreferredSize(new Dimension(getWidth(), 30));

        // Кнопки управления
        JPanel controlButtons = new JPanel(new FlowLayout(FlowLayout.RIGHT, 7, 0));
        controlButtons.setOpaque(false);

        // Кнопка сворачивания (hover - белый)
        minimizeButton = createTitleButton("−", Color.WHITE);
        minimizeButton.setFont(segoeFont);  // Применяем шрифт Segoe UI
        minimizeButton.setFont(minimizeButton.getFont().deriveFont(22f)); // Увеличиваем размер символа

        // Кнопка закрытия (hover - красный)
        closeButton = createTitleButton("×", Color.RED);
        closeButton.setFont(segoeFont);  // Применяем шрифт Segoe UI
        closeButton.setFont(closeButton.getFont().deriveFont(22f)); // Увеличиваем размер символа

        controlButtons.add(minimizeButton);
        controlButtons.add(closeButton);

        titleBar.add(controlButtons, BorderLayout.EAST);

        // Панель для основного контента
        JPanel contentPanel = new JPanel();
        contentPanel.setLayout(new GridBagLayout());
        contentPanel.setBackground(backgroundColorFon);

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Всплывающий список (JComboBox)
        gbc.gridx = 0;
        gbc.gridy = 0;
        JLabel operationLabel = createLabel("Выберите тип обработки:", textColor);
        operationLabel.setFont(segoeFont);  // Применяем шрифт Segoe UI
        contentPanel.add(operationLabel, gbc);

        operationComboBox = new JComboBox<>(new String[]{
                "Обработка AB",
                "Обработка комментариев",
                "Обработка комментариев для Lagerbestellung",
                "Обработка дат поставки"
        });
        operationComboBox.setBackground(backgroundButtonColor);
        operationComboBox.setForeground(textColor);
        operationComboBox.setFont(segoeFont);  // Применяем шрифт Segoe UI

// Убираем белую границу или перекрашиваем в цвет кнопок
        //operationComboBox.setBorder(BorderFactory.createLineBorder(backgroundButtonColor, 1)); // Устанавливаем цвет и толщину границы

// Регулирование размера бокса
        //operationComboBox.setPreferredSize(new Dimension(400, 25)); // Задаем ширину и высоту компонента

// Настройка GridBagConstraints
        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.fill = GridBagConstraints.NONE; // Не растягивать по горизонтали
        gbc.anchor = GridBagConstraints.WEST; // Привязка к левой стороне
        contentPanel.add(operationComboBox, gbc);
        gbc.gridwidth = 1;


        // Поле для ввода пути к файлу и кнопка "Обзор"
        gbc.gridx = 0;
        gbc.gridy = 1;
        JLabel filePathLabel = createLabel("Выберите файл Excel:", textColor);
        filePathLabel.setFont(segoeFont);  // Применяем шрифт Segoe UI
        contentPanel.add(filePathLabel, gbc);

        filePathField = new JTextField();
        filePathField.setBackground(inputColor);
        filePathField.setForeground(textColor);
        filePathField.setFont(segoeFont);  // Применяем шрифт Segoe UI


        gbc.gridx = 1;
        gbc.gridy = 1;
        gbc.weightx = 1.0; // Позволяет JTextField растягиваться по горизонтали
        gbc.fill = GridBagConstraints.HORIZONTAL; // Растягивает компонент на всю доступную ширину
        contentPanel.add(filePathField, gbc);


        browseButton = createRoundedButton("Обзор", backgroundButtonColor, buttonHoverColor, textColor);
        browseButton.setFont(segoeFont);  // Применяем шрифт Segoe UI
        gbc.gridx = 2;
        gbc.weightx = 0;
        browseButton.setPreferredSize(new Dimension(90,25));
        contentPanel.add(browseButton, gbc);

        // Кнопка "Начать обработку"
        startButton = createRoundedButton("Начать обработку", backgroundButtonColor, buttonHoverColor, textColor);
        startButton.setFont(segoeFont);  // Применяем шрифт Segoe UI
        startButton.setPreferredSize(new Dimension(180, 25)); // Увеличение высоты до 50 пикселей
        gbc.gridx = 1;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        contentPanel.add(startButton, gbc);

        // Добавляем панели в основное окно
        add(titleBar, BorderLayout.NORTH);
        add(contentPanel, BorderLayout.CENTER); // Панель основного контента в центре

        // Действия при нажатии на кнопку "Обзор"
        browseButton.addActionListener(e -> {
            FileChoserExample.showFileChooser(this, filePathField); // Вызов JavaFX FileChooser
        });

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Действие при нажатии на кнопку "Начать обработку"
        startButton.addActionListener(e -> {
            int choice = operationComboBox.getSelectedIndex() + 1; // Получаем выбранный тип обработки
            String filePath = filePathField.getText(); // Получаем путь к файлу из текстового поля

            if (filePath.isEmpty()) { // Если путь к файлу не указан
                JOptionPane.showMessageDialog(null, "Пожалуйста, выберите файл Excel.", "Ошибка", JOptionPane.ERROR_MESSAGE); // Показываем сообщение об ошибке
                return;
            }

            try {
                automation.processFile(choice, filePath); // Пытаемся обработать файл
                JOptionPane.showMessageDialog(null, "Обработка завершена.", "Успех", JOptionPane.INFORMATION_MESSAGE); // Показываем сообщение об успешной обработке
            } catch (Exception ex) { // В случае ошибки
                ex.printStackTrace(); // Выводим стек вызовов для отладки
                JOptionPane.showMessageDialog(null, "Произошла ошибка при обработке файла.", "Ошибка", JOptionPane.ERROR_MESSAGE); // Показываем сообщение об ошибке
            }
        });

        // Действие для кнопки сворачивания окна
        minimizeButton.addActionListener(e -> setState(JFrame.ICONIFIED)); // Сворачиваем окно при нажатии

        // Действие для кнопки закрытия окна
        closeButton.addActionListener(e -> System.exit(0)); // Закрываем программу при нажатии

        // Обработка событий для перетаскивания окна за панель заголовка
        titleBar.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                offsetX = e.getX(); // Запоминаем смещение по оси X
                offsetY = e.getY(); // Запоминаем смещение по оси Y
            }
        });

        titleBar.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                int x = e.getXOnScreen(); // Получаем текущую позицию мыши по X
                int y = e.getYOnScreen(); // Получаем текущую позицию мыши по Y
                setLocation(x - offsetX, y - offsetY); // Перемещаем окно в новую позицию
            }
        });
    }

    private int offsetX, offsetY; // Переменные для хранения смещения при перетаскивании окна

    // Метод для создания кнопки с настраиваемыми цветами и скругленными углами
    private JButton createRoundedButton(String text, Color backgroundColor, Color hoverColor, Color textColor) {
        JButton button = new JButton(text) {
            @Override
            protected void paintComponent(Graphics g) {
                // Определяем цвет фона в зависимости от состояния кнопки
                if (getModel().isRollover()) {  // Проверка, наведена ли мышь на кнопку
                    g.setColor(hoverColor);
                } else {
                    g.setColor(backgroundColor);
                }
                g.fillRoundRect(0, 0, getWidth(), getHeight(), 20, 20); // Округленные углы
                super.paintComponent(g);
            }
        };
        button.setForeground(textColor); // Цвет текста
        button.setOpaque(false); // Устанавливаем непрозрачность кнопки
        button.setContentAreaFilled(false); // Отключаем заливку области кнопки
        button.setFocusPainted(false); // Отключаем фокусировку
        button.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20)); // Устанавливаем отступы

        // Обработчики событий для изменения состояния при наведении и уходе мыши
        button.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                button.getModel().setRollover(true); // Устанавливаем состояние наведения
                button.repaint(); // Перерисовываем кнопку
            }

            @Override
            public void mouseExited(MouseEvent e) {
                button.getModel().setRollover(false); // Снимаем состояние наведения
                button.repaint(); // Перерисовываем кнопку
            }
        });

        return button;
    }


    // Метод для создания кнопок заголовка (свернуть/закрыть)
    private JButton createTitleButton(String text, Color hoverColor) {
        JButton button = new JButton(text);
        button.setForeground(Color.GRAY);
        button.setContentAreaFilled(false);
        button.setBorder(new EmptyBorder(0, 0, 0, 10)); // Отступы внутри кнопки
        button.setFocusPainted(false);

        // Обработчики для изменения цвета при наведении
        button.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                button.setForeground(hoverColor);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                button.setForeground(Color.GRAY);
            }
        });

        return button;
    }



    // Метод для создания меток с текстом и настроенным цветом
    private JLabel createLabel(String text, Color textColor) {
        JLabel label = new JLabel(text); // Создаем метку с текстом
        label.setForeground(textColor); // Устанавливаем цвет текста метки
        return label; // Возвращаем созданную метку
    }

    // Главный метод запуска приложения
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            OrderAutomationGUI gui = new OrderAutomationGUI(); // Создаем экземпляр интерфейса
            gui.setVisible(true); // Показываем интерфейс пользователю
        });
    }
}
