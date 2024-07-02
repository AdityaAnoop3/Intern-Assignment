package com.internship.project.TaskReport.Service;

import com.internship.project.TaskReport.DTO.TaskEfforts;

import java.time.LocalDate;
import java.util.List;
import java.util.Set;

public interface ReadExcelService {
    public boolean isSameWeek(LocalDate date1, LocalDate date2);
    public Set<LocalDate> readPublicHolidays(String filepath);
    public List<TaskEfforts> readExcel(String filePath);
}
