module AttendanceLogGenerator {
  requires javafx.controls;
  requires javafx.fxml;
  requires org.apache.poi.ooxml;
  requires javafx.graphics;
  requires org.apache.poi.poi;

  opens com.fhs to javafx.graphics, javafx.fxml;
}
