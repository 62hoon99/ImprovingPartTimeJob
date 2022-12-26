package personal.ImprovingPartTimeJob.controller;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.FileItem;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItem;
import org.springframework.core.io.FileSystemResource;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;
import personal.ImprovingPartTimeJob.service.Service;

import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;
import javax.validation.constraints.NotEmpty;
import java.io.*;
import java.net.MalformedURLException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;

@org.springframework.stereotype.Controller
@Slf4j
@RequiredArgsConstructor
public class Controller {

    private final Service service;

    @PostMapping("/orderFile")
    public void modifyOrderFile(@NotEmpty @RequestParam("orderFile") MultipartFile orderFile, HttpServletResponse response) throws IOException {
        String storedFileName = service.modifyOrderFile(orderFile);
        Workbook wb = WorkbookFactory.create(new File(service.getFileDir() + storedFileName));
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(orderFile.getOriginalFilename(), StandardCharsets.UTF_8));
        wb.write(response.getOutputStream());
        wb.close();

        service.deleteFile(storedFileName);
    }
}
