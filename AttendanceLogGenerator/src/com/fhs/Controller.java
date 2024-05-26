package com.fhs;

import java.io.File;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;

import javafx.collections.FXCollections;
import javafx.concurrent.Task;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.Spinner;
import javafx.scene.control.SpinnerValueFactory;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;

public class Controller {

  @FXML
  private GridPane mainContainer;
  @FXML
  private Button selectFileButton;
  @FXML
  private Button saveFileButton;
  @FXML
  private ComboBox<String> monthComboBox;
  @FXML
  private Spinner<Integer> yearSpinner;
  @FXML
  private Label outputLabel;

  private File selectedFile; // To hold the reference to the selected input file

  @FXML
  public void initialize() {
    initializeDateControls();
    selectFileButton.setOnAction(event -> selectFile());
    saveFileButton.setOnAction(event -> {
      if (selectedFile != null && monthComboBox.getValue() != null
          && yearSpinner.getValue() != null) {
        LocalDate selectedMonth = LocalDate.of(yearSpinner.getValue(),
                                               monthComboBox.getSelectionModel().getSelectedIndex()
                                                                       + 1,
                                               1);
        saveFile(selectedMonth);
      } else {
        outputLabel.setText("Please select both a file and a month/year.");
      }
    });

  }

  private void initializeDateControls() {
    monthComboBox.setItems(FXCollections.observableArrayList("January", "February", "March",
                                                             "April", "May", "June", "July",
                                                             "August", "September", "October",
                                                             "November", "December"));
    LocalDate nextMonth = LocalDate.now().plusMonths(1);
    monthComboBox.getSelectionModel().select(nextMonth.getMonthValue() - 1);

    SpinnerValueFactory.IntegerSpinnerValueFactory yearFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000,
                                                                                                                    2100,
                                                                                                                    nextMonth.getYear());
    yearSpinner.setValueFactory(yearFactory);
  }

  private void selectFile() {
    FileChooser fileChooser = new FileChooser();
    fileChooser.setTitle("Select Name File");
    fileChooser.getExtensionFilters()
               .addAll(new FileChooser.ExtensionFilter("Text Files", "*.txt"));
    selectedFile = fileChooser.showOpenDialog(null);
    if (selectedFile != null) {
      outputLabel.setText("Selected file: " + selectedFile.getName());
    } else {
      outputLabel.setText("No file selected.");
    }
  }

  private void saveFile(LocalDate selectedMonth) {
    FileChooser fileChooser = new FileChooser();
    fileChooser.setTitle("Save Attendance Log");
    fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
    fileChooser.setInitialFileName("Attendance_Log_" + selectedMonth.getMonthValue() + "-"
                                   + selectedMonth.getYear() + ".xlsx");
    File saveFile = fileChooser.showSaveDialog(null);
    if (saveFile != null) {
      generateCalendar(selectedFile, selectedMonth, saveFile);
    } else {
      outputLabel.setText("No save location selected.");
    }
  }

  private void generateCalendar(File inputFile, LocalDate selectedMonth, File outputFile) {
    Task<Void> task = new Task<>() {
      @Override
      protected Void call() throws Exception {
        CalendarGenerator.generate(inputFile, selectedMonth, outputFile);
        return null;
      }

      @Override
      protected void succeeded() {
        super.succeeded();
        outputLabel.setText("Generated calendar for "
                            + selectedMonth.getMonth().getDisplayName(TextStyle.FULL,
                                                                      Locale.getDefault())
                            + " " + selectedMonth.getYear());
      }

      @Override
      protected void failed() {
        super.failed();
        outputLabel.setText("Failed to generate calendar: " + getException().getMessage());
      }
    };
    new Thread(task).start();
  }
}