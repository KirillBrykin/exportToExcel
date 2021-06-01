package ru.magnit.testservice.exportToExcel.model;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.AllArgsConstructor;
import lombok.Data;
import org.springframework.format.annotation.DateTimeFormat;

import java.util.Date;
import java.util.List;

@Data
@AllArgsConstructor
public class WriteOffProcess {
    private String id;
    private String name;
    private String type;
    private String status;
    private Date date;
    private String from;
    private List<WriteOffProcessItem> items;

}
