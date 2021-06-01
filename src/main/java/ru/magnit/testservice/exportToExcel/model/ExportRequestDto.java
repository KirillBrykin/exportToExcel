package ru.magnit.testservice.exportToExcel.model;

import lombok.Data;

import java.util.List;

@Data
public class ExportRequestDto {
    private String processId;
    private List<String> exportFields;
}
