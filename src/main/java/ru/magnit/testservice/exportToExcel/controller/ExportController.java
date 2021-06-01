package ru.magnit.testservice.exportToExcel.controller;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;
import ru.magnit.testservice.exportToExcel.model.ExportRequestDto;
import ru.magnit.testservice.exportToExcel.model.WriteOffProcess;
import ru.magnit.testservice.exportToExcel.model.WriteOffProcessItem;
import ru.magnit.testservice.exportToExcel.service.ExportService;

import java.util.Date;

@RestController
@RequiredArgsConstructor
public class ExportController {

    private final ExportService exportService;

    @GetMapping("export")
    public ResponseEntity<StreamingResponseBody> export(@RequestBody ExportRequestDto exportRequestDto) {
        System.out.println("ExportController.export");
        System.out.println("exportRequestDto = " + exportRequestDto);

        Workbook workbook = exportService.generateWorkbook(exportRequestDto);

        return ResponseEntity
                .ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, "inline;filename=\"Write-off-process.xlsx\"")
                .body(workbook::write);
    }


}
