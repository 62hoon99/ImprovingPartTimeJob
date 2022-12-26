package personal.ImprovingPartTimeJob.repository;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.tomcat.util.http.fileupload.FileItem;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItem;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.DefaultResourceLoader;
import org.springframework.core.io.Resource;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.*;
import java.net.URI;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.UUID;

@org.springframework.stereotype.Repository
@RequiredArgsConstructor
@Slf4j
public class Repository {

    @Value("${spring.file-dir}")
    private String rootDir;

    @Value("${file.password}")
    private String password;

    public String modifyOrderFile(MultipartFile file) throws IOException {
        UUID identifier = UUID.randomUUID();
        String fileDirAndName = getFileDir() + identifier + ".xlsx";
        file.transferTo(new File(fileDirAndName));
        saveModifiedOrderFile(fileDirAndName);
        return identifier + ".xlsx";
    }

    private void saveModifiedOrderFile(String fileDirAndName) throws IOException {
        Workbook workbook = WorkbookFactory.create(new File(fileDirAndName), password);
        Sheet sheet = workbook.getSheetAt(0);
        sheet.shiftRows(1, sheet.getLastRowNum(), -1);
        FileOutputStream fileOutputStream = new FileOutputStream(fileDirAndName);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    public String getFileDir() {
        return System.getProperty("user.dir") + this.rootDir;
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
