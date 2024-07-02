package com.internship.project.TaskReport.Controller;

import com.internship.project.TaskReport.DTO.TaskEfforts;
import com.internship.project.TaskReport.Service.GenerateReportService;
import com.internship.project.TaskReport.Service.ReadExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import java.util.List;


@RestController
public class TaskReportController {
    @Autowired
    ReadExcelService readExcelService;
    @Autowired
    GenerateReportService generateReportService;
    @Value("${filepath.input}")
    private String filepathInput;

    @GetMapping("/ControlMethod")
    public void ControllerMethod() {
        List<TaskEfforts> taskEfforts = readExcelService.readExcel(filepathInput);
        generateReportService.generateReport(taskEfforts);
    }
}
