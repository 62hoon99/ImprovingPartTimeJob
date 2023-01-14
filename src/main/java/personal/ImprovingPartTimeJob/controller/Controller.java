package personal.ImprovingPartTimeJob.controller;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;
import personal.ImprovingPartTimeJob.controller.form.OrderFileForm;
import personal.ImprovingPartTimeJob.service.Service;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Objects;

@org.springframework.stereotype.Controller
@Slf4j
@RequiredArgsConstructor
public class Controller {

    @Value("${file.uploadFileName}")
    private String uploadFileName;

    @Value("${file.password}")
    private String password;

    private final Service service;

    @GetMapping("/")
    public String home(Model model) {
        model.addAttribute("orderFileForm", new OrderFileForm(password));
        return "home";
    }

    @PostMapping("/orderFile")
    public String modifyOrderFile(OrderFileForm orderFileForm, RedirectAttributes redirectAttributes) {
        MultipartFile orderFile = orderFileForm.getOrderFile();
        String password = orderFileForm.getPassword();

        if(orderFile.isEmpty()) {
            return "redirect:/";
        }

        try {
            String storedFileName = service.modifyOrderFile(orderFile, password);
            redirectAttributes.addAttribute("storedFileName", storedFileName);
            redirectAttributes.addAttribute("fileName", orderFile.getOriginalFilename());
            return "redirect:/download";
        } catch (Exception e) {
            log.error(e.getMessage());
            return "redirect:/";
        }
    }

    @GetMapping("/download")
    public void downloadOrderFile(String storedFileName, String fileName,
                                  HttpServletResponse response) {
        try {
            containFileToResponse(storedFileName, response, fileName);
            service.deleteFile(storedFileName);
        } catch (Exception e) {
            log.error(e.getMessage());
        }
    }

    @PostMapping("/batchFile")
    public String createBatchFile(@RequestParam MultipartFile orderFile,
                                @RequestParam MultipartFile receiptFile,
                                @RequestParam String turn, RedirectAttributes redirectAttributes) {
        if(orderFile.isEmpty() || receiptFile.isEmpty() || turn.isEmpty()) {
            return "redirect:/";
        }

        try {
            String storedFileName = service.createBatchFile(orderFile, receiptFile);
            redirectAttributes.addAttribute("storedFileName", storedFileName);
            redirectAttributes.addAttribute("fileName", getBatchFileName(turn));
            return "redirect:/download";
        } catch (Exception e) {
            log.error(e.getMessage());
            return "redirect:/";
        }
    }

    @PostMapping("/cleanup")
    public String cleanupDirectory() {
        service.cleanUpDirectory();
        return "redirect:/";
    }

    private String getBatchFileName(String turn) {
        return getDateFormat() + getTurnFormat(turn) + uploadFileName;
    }

    private String getTurnFormat(String turn) {
        return turn + "차_";
    }

    private String getDateFormat() {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy_MM_dd_");
        return LocalDate.now().format(formatter);
    }

    private void containFileToResponse(String storedFileName, HttpServletResponse response, String fileName) throws IOException {
        Workbook wb = getWorkbook(storedFileName);
        response.setContentType("ms-vnd/excel");
        response.setHeader("Content-Disposition", "attachment;filename=" +
                URLEncoder.encode(Objects.requireNonNull(fileName), StandardCharsets.UTF_8));
        wb.write(response.getOutputStream());
        wb.close();
    }

    private Workbook getWorkbook(String storedFileName) throws IOException {
        return WorkbookFactory.create(new File(service.getFileDir() + storedFileName));
    }
}
