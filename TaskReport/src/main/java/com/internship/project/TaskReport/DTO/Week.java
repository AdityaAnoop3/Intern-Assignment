package com.internship.project.TaskReport.DTO;

import java.time.LocalDate;

public class Week {
    private int weekOfYear;
    private LocalDate startDate;
    private LocalDate endDate;
    private LocalDate nextStartDate;

    public Week() {
    }

    public int getWeekOfYear() {
        return weekOfYear;
    }

    public void setWeekOfYear(int weekOfYear) {
        this.weekOfYear = weekOfYear;
    }

    public LocalDate getStartDate() {
        return startDate;
    }

    public void setStartDate(LocalDate startDate) {
        this.startDate = startDate;
    }

    public LocalDate getEndDate() {
        return endDate;
    }

    public void setEndDate(LocalDate endDate) {
        this.endDate = endDate;
    }

    public LocalDate getNextStartDate() {
        return nextStartDate;
    }

    public void setNextStartDate(LocalDate nextStartDate) {
        this.nextStartDate = nextStartDate;
    }

}
