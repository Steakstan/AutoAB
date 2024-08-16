module AutoAB {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.graphics;
    requires javafx.base;

    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.ooxml.schemas;
    requires org.apache.commons.compress;
    requires org.apache.commons.lang3;
    requires org.apache.xmlbeans;

    requires com.sun.jna;
    requires com.sun.jna.platform;
    requires java.logging;
    requires java.desktop;

    opens org.example to javafx.fxml;
    exports org.example;
}
