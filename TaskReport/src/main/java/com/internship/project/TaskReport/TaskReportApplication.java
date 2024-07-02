package com.internship.project.TaskReport;

import com.internship.project.TaskReport.Controller.TaskReportController;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class TaskReportApplication implements CommandLineRunner {

	@Autowired
	private TaskReportController taskReportController;

	public static void main(String[] args) {
		SpringApplication.run(TaskReportApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		taskReportController.ControllerMethod();
	}
}
