package personal.ImprovingPartTimeJob.controller;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import personal.ImprovingPartTimeJob.service.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.Objects;

@org.springframework.stereotype.Controller
@Slf4j
@RequiredArgsConstructor
public class Controller {

    private final Service service;

    @PostMapping("/orderFile")
    public void modifyOrderFile(@RequestParam MultipartFile orderFile,
                                HttpServletResponse response) throws IOException {
        String storedFileName = service.modifyOrderFile(orderFile);
        containFileToResponse(storedFileName, response, orderFile.getOriginalFilename());
        service.deleteFile(storedFileName);
    }

    @PostMapping("/batchFile")
    public void createBatchFile(@RequestParam MultipartFile orderFile,
                                @RequestParam MultipartFile receiptFile,
                                @RequestParam String turn,
                                HttpServletResponse response) throws IOException {
        String storedBatchFileName = service.createBatchFile(orderFile, receiptFile, turn);
        containFileToResponse(storedBatchFileName, response, storedBatchFileName);
        service.deleteFile(storedBatchFileName);
    }

    private void containFileToResponse(String storedFileName, HttpServletResponse response, String orderFile) throws IOException {
        Workbook wb = getWorkbook(storedFileName);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(Objects.requireNonNull(orderFile), StandardCharsets.UTF_8));
        wb.write(response.getOutputStream());
        wb.close();
    }

    private Workbook getWorkbook(String storedFileName) throws IOException {
        return WorkbookFactory.create(new File(service.getFileDir() + storedFileName));
    }
}
