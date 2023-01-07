package personal.ImprovingPartTimeJob.controller;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import personal.ImprovingPartTimeJob.service.Service;

import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.NotEmpty;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Objects;

@org.springframework.stereotype.Controller
@Slf4j
@RequiredArgsConstructor
public class Controller {

    private final Service service;

    @PostMapping("/orderFile")
    public void modifyOrderFile(@NotEmpty @RequestParam("orderFile") MultipartFile orderFile, HttpServletResponse response) throws IOException {
        String storedFileName = service.modifyOrderFile(orderFile);
        Workbook wb = getWorkbook(storedFileName);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(Objects.requireNonNull(orderFile.getOriginalFilename()), StandardCharsets.UTF_8));
        wb.write(response.getOutputStream());
        wb.close();

        service.deleteFile(storedFileName);
    }

    @PostMapping("/batchFile")
    public void createBatchFile(@RequestParam MultipartFile orderFile,
                                @RequestParam MultipartFile receiptFile,
                                HttpServletResponse response) throws IOException {

        String storedBatchFileName = service.createBatchFile(orderFile, receiptFile);
        Workbook wb = getWorkbook(storedBatchFileName);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(Objects.requireNonNull(storedBatchFileName), StandardCharsets.UTF_8));
        wb.write(response.getOutputStream());
        wb.close();

        service.deleteFile(storedBatchFileName);
    }

    private Workbook getWorkbook(String storedFileName) throws IOException {
        Workbook wb = WorkbookFactory.create(new File(service.getFileDir() + storedFileName));
        return wb;
    }

    @Data
    private class CreateBatchFileForm {
        private MultipartFile orderFile;
        private MultipartFile receiptFile;
    }
}