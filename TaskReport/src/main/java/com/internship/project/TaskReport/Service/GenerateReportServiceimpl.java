package com.internship.project.TaskReport.Service;

//Import BLock

import com.internship.project.TaskReport.DTO.TaskEfforts;
import com.internship.project.TaskReport.DTO.Week;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

//Service Class Block
@Service
public class GenerateReportServiceimpl implements GenerateReportService {
    @Value("${prepared.by}")
    private String preparedBy;

    @Value("${filepath.output}")
    private String filepathOutput;


    public void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            if (cell == null) {
                cell = table.getRow(row).addNewTableCell();
            }
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            if (colIndex == fromCol) {
                // The first merged cell is set with RESTART merge value
                tcPr.addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                tcPr.addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


    public void generateReport(List<TaskEfforts> taskEfforts) {
        // Group the task efforts by task and week
        Map<String, Map<Week, List<TaskEfforts>>> taskEffortsByTaskWeek = taskEfforts.stream()
                .collect(Collectors.groupingBy(TaskEfforts::getTask,
                        Collectors.groupingBy(TaskEfforts::getWeek)));

        // Iterate over each task
        for (String task : taskEffortsByTaskWeek.keySet()) {
            XWPFDocument document = new XWPFDocument();
            XWPFParagraph title = document.createParagraph();
            XWPFRun run = title.createRun();
            run.setText("Actual Task Tracking - " + task);
            run.setBold(true);
            run.setFontSize(24);

            XWPFTable table = document.createTable();

            // Create table header
            XWPFTableRow headerRow = table.getRow(0);
            headerRow.getCell(0).setText("SN");
            headerRow.addNewTableCell().setText("Week of the year");
            headerRow.addNewTableCell().setText("Week Start Date");
            headerRow.addNewTableCell().setText("Week end Date");
            headerRow.addNewTableCell().setText("Actual effort");
            headerRow.addNewTableCell().setText("Date");
            headerRow.addNewTableCell().setText("Prepared By");

            // Sort the weeks based on the weekOfYear attribute
            List<Week> sortedWeeks = new ArrayList<>(taskEffortsByTaskWeek.get(task).keySet());
            sortedWeeks.sort(Comparator.comparing(Week::getWeekOfYear));

            int totalEffort = 0;

            // Iterate over each week
            int snCounter = 0;
            for (Week week : sortedWeeks) {
                List<TaskEfforts> weekEfforts = taskEffortsByTaskWeek.get(task).get(week);

                // Fill table with data
                snCounter++;
                XWPFTableRow row = table.createRow();
                row.getCell(0).setText(String.valueOf(snCounter));
                row.getCell(1).setText(String.valueOf(week.getWeekOfYear()));
                row.getCell(2).setText(week.getStartDate().toString());
                row.getCell(3).setText(week.getEndDate().toString());

                // Create a new paragraph in the 'Actual effort' cell
                XWPFParagraph actualEffortParagraph = row.getCell(4).addParagraph();
                actualEffortParagraph.setSpacingAfter(0);  // Remove extra spacing after the paragraph

                // Iterate over each task effort in the week
                for (TaskEfforts taskEffort : weekEfforts) {
                    // Create a new run in the 'Actual effort' cell for each task effort
                    XWPFRun actualEffortRun = actualEffortParagraph.createRun();
                    actualEffortRun.setText(taskEffort.getEffort() + " (" + taskEffort.getPerson() + ")");
                    actualEffortRun.addBreak();  // Add a line break after each task effort
                    totalEffort += taskEffort.getEffort();
                }

                row.getCell(5).setText(week.getNextStartDate().toString());
                row.getCell(6).setText(preparedBy);
            }

            // Add a total row at the end of the table
            XWPFTableRow totalRow = table.createRow();

            // Remove the cell that is automatically created with the new row
            totalRow.removeCell(0);

            // Merge all cells in the total row
            int numberOfColumns = table.getRow(0).getTableCells().size();
            int totalRowIndex = table.getRows().indexOf(totalRow);
            mergeCellsHorizontally(table, totalRowIndex, 0, numberOfColumns - 1);

            // Add the total cell
            XWPFTableCell totalCell = totalRow.getCell(0);

            // Create a new paragraph in the total cell
            XWPFParagraph totalParagraph = totalCell.addParagraph();

            // Set the alignment of the paragraph to CENTER
            totalParagraph.setAlignment(ParagraphAlignment.CENTER);
            totalParagraph.setVerticalAlignment(TextAlignment.CENTER);

            // Add the total effort text to the paragraph
            XWPFRun totalRun = totalParagraph.createRun();
            totalRun.setText("Total: " + totalEffort);

            // Save document
            try (FileOutputStream out = new FileOutputStream(filepathOutput + task + ".docx")) {
                document.write(out);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}