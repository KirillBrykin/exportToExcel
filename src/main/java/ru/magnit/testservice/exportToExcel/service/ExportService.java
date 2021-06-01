package ru.magnit.testservice.exportToExcel.service;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import ru.magnit.testservice.exportToExcel.model.ExportRequestDto;
import ru.magnit.testservice.exportToExcel.model.WriteOffProcess;
import ru.magnit.testservice.exportToExcel.model.WriteOffProcessItem;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Service
public class ExportService {

    public Workbook generateWorkbook(ExportRequestDto exportRequestDto) {
        Workbook workbook = null;
        WriteOffProcess writeOffProcess = getData(exportRequestDto.getProcessId());
        if (writeOffProcess != null) {
            workbook = new XSSFWorkbook();

//            RegionUtil.setBorderBottom(CellStyle.BORDER_MEDIUM, region, sheet, wb);




            Sheet sheet = workbook.createSheet("write-off-process");
            CellStyle headerCellStyle = createHeaderCellStyle(workbook);
            createTableHeader(sheet, writeOffProcess, headerCellStyle, workbook);
            createTable(sheet, writeOffProcess);
        }
        return workbook;
    }

    private CellStyle createHeaderCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
//        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
//        cellStyle.setBorderTop(BorderStyle.MEDIUM);
//        cellStyle.setBorderRight(BorderStyle.MEDIUM);
//        cellStyle.setBorderLeft(BorderStyle.MEDIUM);
        return cellStyle;
    }

    private CellStyle setBoldFont (Workbook workbook, CellStyle cellStyle) {
        Font font = workbook.createFont();
        font.setBold(true);
        cellStyle.setFont(font);
        return cellStyle;
    }



    private void createTableHeader(Sheet sheet, WriteOffProcess writeOffProcess, CellStyle headerCellStyle, Workbook workbook) {
        Row header0 = sheet.createRow(0);
        Cell headerCell00 = header0.createCell(0);
        headerCell00.setCellValue("Операция");
        headerCell00.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell01 = header0.createCell(1);
        headerCell01.setCellValue(writeOffProcess.getName());
        Cell headerCell03 = header0.createCell(3);
        headerCell03.setCellValue("Статус");
        headerCell03.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell04 = header0.createCell(4);
        headerCell04.setCellValue(writeOffProcess.getStatus());

        Row header1 = sheet.createRow(1);
        Cell headerCell10 = header1.createCell(0);
        headerCell10.setCellValue("Номер");
        headerCell10.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell11 = header1.createCell(1);
        headerCell11.setCellValue(writeOffProcess.getId());
        Cell headerCell13 = header1.createCell(3);
        headerCell13.setCellValue("Дата");
        headerCell13.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell14 = header1.createCell(4);
        headerCell14.setCellValue(new SimpleDateFormat("dd.MM.yyyy").format(writeOffProcess.getDate()));

        Row header2 = sheet.createRow(2);
        Cell headerCell20 = header2.createCell(0);
        headerCell20.setCellValue("Тип");
        headerCell20.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell21 = header2.createCell(1);
        headerCell21.setCellValue(writeOffProcess.getType());
        Cell headerCell23 = header2.createCell(3);
        headerCell23.setCellValue("МХ(откуда)");
        headerCell23.setCellStyle(setBoldFont(workbook, headerCellStyle));
        Cell headerCell24 = header2.createCell(4);
        headerCell24.setCellValue(writeOffProcess.getFrom());

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
    }


    private void createTable(Sheet sheet, WriteOffProcess writeOffProcess) {
        int nullRow = 4;

        Row headerRow = sheet.createRow(nullRow++);
        Cell headerCell1 = headerRow.createCell(0);
        headerCell1.setCellValue("Код товара");
        sheet.autoSizeColumn(0);
        Cell headerCell2 = headerRow.createCell(1);
        headerCell2.setCellValue("Наименование товара");
        sheet.autoSizeColumn(1);
        Cell headerCell3 = headerRow.createCell(2);
        headerCell3.setCellValue("ед.изм.");
        sheet.autoSizeColumn(2);
        Cell headerCell4 = headerRow.createCell(3);
        headerCell4.setCellValue("Количество");
        sheet.autoSizeColumn(3);

        for (int i = 0; i < writeOffProcess.getItems().size(); i++) {
            Row row = sheet.createRow(nullRow++);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(writeOffProcess.getItems().get(i).getCode());
            sheet.autoSizeColumn(0);
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(writeOffProcess.getItems().get(i).getName());
            sheet.autoSizeColumn(1);
            Cell cell3 = row.createCell(2);
            cell3.setCellValue(writeOffProcess.getItems().get(i).getMeasureUnit());
            sheet.autoSizeColumn(2);
            Cell cell4 = row.createCell(3);
            cell4.setCellValue(writeOffProcess.getItems().get(i).getCount());
            sheet.autoSizeColumn(3);
        }

    }


    private WriteOffProcess getData(String processId) {
        WriteOffProcessItem writeOffProcessItem1 = new WriteOffProcessItem("name1", "code1", "measureUnit1", "count1");
        WriteOffProcessItem writeOffProcessItem2 = new WriteOffProcessItem("name2", "code2", "measureUnit2", "count2");
        WriteOffProcessItem writeOffProcessItem3 = new WriteOffProcessItem("name3", "code3", "measureUnit3", "count3");
        List<WriteOffProcessItem> items = new ArrayList<>();
        items.add(writeOffProcessItem1);
        items.add(writeOffProcessItem2);
        items.add(writeOffProcessItem3);
        WriteOffProcess writeOffProcess = new WriteOffProcess(processId, "name1", "type1", "status1", new Date(), "from1", items);
        return writeOffProcess;
    }

}
