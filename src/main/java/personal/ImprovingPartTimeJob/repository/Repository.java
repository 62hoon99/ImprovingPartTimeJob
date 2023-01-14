package personal.ImprovingPartTimeJob.repository;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Objects;
import java.util.UUID;

@org.springframework.stereotype.Repository
@RequiredArgsConstructor
@Slf4j
public class Repository {

    @Value("${spring.file-dir}")
    private String rootDir;

    @Value("${file.shippingMethod}")
    private String shippingMethod;

    @Value("${file.shippingCompany}")
    private String shippingCompany;

    @Value("${file.uploadFileName}")
    private String uploadFileName;

    @Value("${spring.fixFile-dir}")
    private String fixFileDir;

    public void saveModifiedOrderFile(String storedFileName, String password) throws IOException {
        Workbook workbook = WorkbookFactory.create(new File(getFileDir() + storedFileName), password);
        Sheet sheet = workbook.getSheetAt(0);
        sheet.shiftRows(1, sheet.getLastRowNum(), -1);
        FileOutputStream fileOutputStream = new FileOutputStream(getFileDir() + storedFileName);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    public void cleanupDirectory() {
        File file = new File(getFileDir());

        for (String fileName : Objects.requireNonNull(file.list())) {
            deleteFile(fileName);
        }
    }

    public String saveFile(MultipartFile file) throws IOException {
        String storedFileName = getRandomStoredFileName(extractExt(Objects.requireNonNull(file.getOriginalFilename())));
        file.transferTo(new File(getFileDir() + storedFileName));
        return storedFileName;
    }

    private String getRandomStoredFileName(String ext) {
        return UUID.randomUUID() + ext;
    }

    private String extractExt(String originalFilename) {
        int pos = originalFilename.lastIndexOf(".");
        return originalFilename.substring(pos);
    }

    public String getFileDir() {
        return System.getProperty("user.dir") + rootDir;
    }

    public String getFixFileDir() {
        return System.getProperty("user.dir") + fixFileDir;
    }

    public String getUploadFileName() {
        return uploadFileName;
    }

    public String getShippingMethod() {
        return shippingMethod;
    }

    public String getShippingCompany() {
        return shippingCompany;
    }

    public void deleteFile(String storedFileName) {
        File file = new File(getFileDir() + storedFileName);
        if (file.delete()) {
            log.info("File delete success");
            return;
        }
        log.info("File delete fail");
    }
}
