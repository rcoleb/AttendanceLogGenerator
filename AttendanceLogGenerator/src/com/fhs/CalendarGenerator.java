package com.fhs;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CalendarGenerator {
  public static void generate(File nameFile, LocalDate monthYear, File saveFile) throws Exception {
    List<String> names = Files.readAllLines(Paths.get(nameFile.toURI()));
    try (Workbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Attendance Log");
      createTitle(sheet, monthYear);
      generateHeaders(sheet, monthYear);
      fillNamesAndDates(sheet, names, monthYear);
      styleSheet(sheet);

      // Save the file
      try (FileOutputStream out = new FileOutputStream(saveFile)) {
        workbook.write(out);
      }
    }
  }

  private static void createTitle(Sheet sheet, LocalDate monthYear) {
    Row titleRow = sheet.createRow(0);
    Cell titleCell = titleRow.createCell(0);
    titleCell.setCellValue("Members Practice Log - "
                           + monthYear.getMonth().getDisplayName(TextStyle.FULL,
                                                                 Locale.getDefault())
                           + " " + monthYear.getYear());
    int daysInMonth = monthYear.lengthOfMonth();
    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, daysInMonth));
    CellStyle style = sheet.getWorkbook().createCellStyle();
    style.setAlignment(HorizontalAlignment.CENTER);
    titleCell.setCellStyle(style);
  }

  private static void generateHeaders(Sheet sheet, LocalDate monthYear) {
    Row headerRow = sheet.createRow(1);
    CellStyle headerStyle = sheet.getWorkbook().createCellStyle();
    headerStyle.setAlignment(HorizontalAlignment.CENTER);
    LocalDate firstDayOfMonth = monthYear.withDayOfMonth(1);
    DayOfWeek currentDayOfWeek = firstDayOfMonth.getDayOfWeek();
    int daysInMonth = monthYear.lengthOfMonth();

    for (int i = 0; i < daysInMonth; i++) { // Assuming a week has 7 days
      Cell cell = headerRow.createCell(i + 1);
      cell.setCellValue(Character.toString(currentDayOfWeek.getDisplayName(TextStyle.SHORT,
                                                                           Locale.getDefault())
                                                           .charAt(0)));
      cell.setCellStyle(headerStyle);
      currentDayOfWeek = currentDayOfWeek.plus(1);
    }
  }

  private static void fillNamesAndDates(Sheet sheet, List<String> names, LocalDate monthYear) {
    LocalDate firstDayOfMonth = monthYear.withDayOfMonth(1);
    int daysInMonth = monthYear.lengthOfMonth();

    int rowNum = 2;
    for (String name : names) {
      Row row = sheet.createRow(rowNum);
      Cell nameCell = row.createCell(0);
      nameCell.setCellValue(name);
      for (int i = 0; i < daysInMonth; i++) {
        LocalDate date = firstDayOfMonth.plusDays(i);
        if (date.getMonth() == monthYear.getMonth()) {
          Cell cell = row.createCell(i + 1);
          cell.setCellValue(i + 1);
        }
      }
      rowNum++;
    }
  }

  static class StyleBuilder {

    private String fontName = null;
    private short fontHeightInPoints = 0;
    private IndexedColors foreColor = null;
    private HorizontalAlignment alignment = null;

    public StyleBuilder setFontName(String fontName) {
      this.fontName = fontName;
      return this;
    }

    public StyleBuilder setFontHeightInPoints(short fontHeightInPoints) {
      this.fontHeightInPoints = fontHeightInPoints;
      return this;
    }

    public StyleBuilder setForegroundColor(IndexedColors color) {
      this.foreColor = color;
      return this;
    }

    public StyleBuilder setAlignment(HorizontalAlignment alignment) {
      this.alignment = alignment;
      return this;
    }

    public CellStyle build(Sheet sheet) {
      CellStyle style = sheet.getWorkbook().createCellStyle();
      Font font = sheet.getWorkbook().createFont();
      if (fontName != null) {
        font.setFontName(fontName);
      }
      if (fontHeightInPoints != 0) {
        font.setFontHeightInPoints(fontHeightInPoints);
      }
      if (foreColor != null) {
        style.setFillForegroundColor(foreColor.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
      }
      if (alignment != null) {
        style.setAlignment(alignment);
      }
      style.setFont(font);
      return style;
    }

  }

  private static void styleSheet(Sheet sheet) {
    CellStyle altStyle1Date = new StyleBuilder().setAlignment(HorizontalAlignment.CENTER)
                                                .setFontHeightInPoints((short) 10)
                                                .setFontName("Helvetica")
                                                .setForegroundColor(IndexedColors.LIGHT_GREEN)
                                                .build(sheet);
    CellStyle altStyle2Date = new StyleBuilder().setAlignment(HorizontalAlignment.CENTER)
                                                .setFontHeightInPoints((short) 10)
                                                .setFontName("Helvetica").build(sheet);
    CellStyle altStyle1Name = new StyleBuilder().setAlignment(HorizontalAlignment.LEFT)
                                                .setFontHeightInPoints((short) 10)
                                                .setFontName("Helvetica")
                                                .setForegroundColor(IndexedColors.LIGHT_GREEN)
                                                .build(sheet);
    CellStyle altStyle2Name = new StyleBuilder().setAlignment(HorizontalAlignment.LEFT)
                                                .setFontHeightInPoints((short) 10)
                                                .setFontName("Helvetica").build(sheet);

    Row row = sheet.getRow(0);
    for (Cell cell : row) {
      cell.setCellStyle(altStyle1Date);
    }

    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      row = sheet.getRow(i);

      if (row != null) {
        Iterator<Cell> cellIter = row.cellIterator();
        cellIter.next().setCellStyle(i % 2 == 0 ? altStyle1Name : altStyle2Name);
        int ind = i;
        cellIter.forEachRemaining(cell -> {
          cell.setCellStyle(ind % 2 == 0 ? altStyle1Date : altStyle2Date);
        });
      }
    }
  }

}
