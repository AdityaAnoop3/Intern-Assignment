package com.internship.project.TaskReport.Service;

//Import BLock

import com.internship.project.TaskReport.DTO.TaskEfforts;
import com.internship.project.TaskReport.DTO.Week;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoField;
import java.time.temporal.WeekFields;
import java.util.*;

//Service Class Block
@Service
public class ReadExcelServiceimpl implements ReadExcelService {

    @Value("${csv.input}")
    private String csvInput;

    public boolean isSameWeek(LocalDate date1, LocalDate date2) {
        WeekFields weekFields = WeekFields.of(Locale.getDefault());
        return date1.get(weekFields.weekOfWeekBasedYear()) == date2.get(weekFields.weekOfWeekBasedYear()) &&
                date1.get(ChronoField.YEAR) == date2.get(ChronoField.YEAR);
    }

    public Set<LocalDate> readPublicHolidays(String filepath) {
        Set<LocalDate> publicHolidays = new HashSet<>();

        try (BufferedReader reader = Files.newBufferedReader(Paths.get(filepath))) {
            String line;
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");

            while ((line = reader.readLine()) != null) {
                line = line.replace("\uFEFF", ""); // remove BOM
                LocalDate holiday = LocalDate.parse(line, formatter);
                publicHolidays.add(holiday);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return publicHolidays;
    }

    public List<TaskEfforts> readExcel(String filePath) {
        List<TaskEfforts> taskEfforts = new ArrayList<>();
        Set<LocalDate> publicHolidays = readPublicHolidays(csvInput);
        try {
            FileInputStream excelFile = new FileInputStream(filePath);
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Get the headers (person's names)
            Row headerRow = iterator.next();
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            Week[] currentWeek = new Week[1];

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();

                // Process the date cell
                Cell dateCell = currentRow.getCell(0);
                if (DateUtil.isCellDateFormatted(dateCell)) {
                    Date dateInCell = dateCell.getDateCellValue();
                    LocalDate date = dateInCell.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();

                    if (publicHolidays.contains(date)) {
                        continue;
                    }

                    Week week = new Week();
                    week.setWeekOfYear(date.get(WeekFields.of(Locale.getDefault()).weekOfWeekBasedYear()));
                    week.setStartDate(date.with(DayOfWeek.MONDAY));
                    week.setEndDate(date.with(DayOfWeek.FRIDAY));
                    week.setNextStartDate(date.plusWeeks(1).with(DayOfWeek.MONDAY));

                    if (currentWeek[0] == null || !isSameWeek(currentWeek[0].getStartDate(), date)) {
                        currentWeek[0] = new Week();
                        currentWeek[0].setWeekOfYear(date.get(WeekFields.of(Locale.getDefault()).weekOfWeekBasedYear()));
                        currentWeek[0].setStartDate(date.with(DayOfWeek.MONDAY));
                        currentWeek[0].setEndDate(date.with(DayOfWeek.FRIDAY));
                        currentWeek[0].setNextStartDate(date.plusWeeks(1).with(DayOfWeek.MONDAY));
                    }

                    // Iterate over the rest of the cells in the row (the tasks)
                    for (int i = 1; i < currentRow.getPhysicalNumberOfCells(); i++) {
                        Cell taskCell = currentRow.getCell(i);
                        if (taskCell.getCellType() == CellType.STRING) {
                            String task = taskCell.getStringCellValue();
                            String person = headers.get(i);

                            // Find the TaskEfforts object for the current task, week, and person
                            TaskEfforts taskEffort = null;
                            for (TaskEfforts te : taskEfforts) {
                                if (te.getTask().equals(task) && te.getWeek().equals(currentWeek[0]) && te.getPerson().equals(person)) {
                                    taskEffort = te;
                                    break;
                                }
                            }

                            // If the TaskEfforts object does not exist, create a new one
                            if (taskEffort == null) {
                                taskEffort = new TaskEfforts();
                                taskEffort.setTask(task);
                                taskEffort.setWeek(currentWeek[0]);
                                taskEffort.setPerson(person);
                                taskEffort.setEffort(0);
                                taskEfforts.add(taskEffort);
                            }

                            // Increment the effort for the current TaskEfforts object
                            taskEffort.setEffort(taskEffort.getEffort() + 1);
                        }
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return taskEfforts;
    }


}