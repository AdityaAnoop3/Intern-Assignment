package com.internship.project.TaskReport.Service;

import com.internship.project.TaskReport.DTO.TaskEfforts;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.List;

public interface GenerateReportService {
   public void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol);
   public void generateReport(List<TaskEfforts> taskEfforts);
}
