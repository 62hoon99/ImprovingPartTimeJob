package personal.ImprovingPartTimeJob.repository;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.tomcat.jni.FileInfo;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@org.springframework.stereotype.Repository
@RequiredArgsConstructor
@Slf4j
public class Repository {

    @Value("${spring.file-dir}")
    private String rootDir;

    @Value("${file.password}")
    private String password;

    @Value("${file.shippingMethod}")
    private String shippingMethod;

    @Value("${file.shippingCompany}")
    private String shippingCompany;

    @Value("${file.uploadFileName}")
    private String uploadFileName;

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

    public String createBatchFile(MultipartFile orderFile, MultipartFile receiptFile) throws IOException {
        String orderFileName = saveFile(orderFile);
        String receiptFileName = saveFile(receiptFile);

        Workbook wbBatchFile = WorkbookFactory.create(new File(getFileDir() + uploadFileName));
        Workbook wbOrderFile = WorkbookFactory.create(new File(getFileDir() + orderFileName));
        Workbook wbReceiptFile = WorkbookFactory.create(new File(getFileDir() + receiptFileName));

        Sheet sheetBatch = wbBatchFile.getSheetAt(0);
        Sheet sheetOrder = wbOrderFile.getSheetAt(0);
        Sheet sheetReceipt = wbReceiptFile.getSheetAt(0);

        List<String> orderNumber = getColumnValues(sheetOrder, 0);
        List<String> waybillNumber = getColumnValues(sheetReceipt, 5);// F열
        List<String> payeeName = getColumnValues(sheetReceipt, 10);

        setBatchFileValues(sheetBatch, orderNumber, waybillNumber, payeeName);

        String storedBatchFileName = "날짜" + uploadFileName;
        FileOutputStream fileOutputStream = new FileOutputStream(getFileDir() + storedBatchFileName);
        wbBatchFile.write(fileOutputStream);
        fileOutputStream.close();

        wbBatchFile.close();
        wbOrderFile.close();
        wbReceiptFile.close();

        deleteFile(orderFileName);
        deleteFile(receiptFileName);

        return storedBatchFileName;
    }

    private String saveFile(MultipartFile file) throws IOException {
        UUID identifier = UUID.randomUUID();
        String storedFileName = identifier + extractExt(file.getOriginalFilename());
        file.transferTo(new File(getFileDir() + storedFileName));
        return storedFileName;
    }

    private List<String> getColumnValues(Sheet sheetOrder, int index) {
        List<String> valueList = new ArrayList<>();
        for(int i = 1; i <= sheetOrder.getLastRowNum(); i++) {
            valueList.add(sheetOrder.getRow(i).getCell(index).getStringCellValue());
        }
        return valueList;
    }

    private void setBatchFileValues(Sheet sheetUpload, List<String> orderNumber, List<String> waybillNumber, List<String> recieverName) {
        for(int i = 0; i < orderNumber.size(); i++) {
            Row row = sheetUpload.createRow(i + 1);
            row.createCell(0).setCellValue(orderNumber.get(i));
            row.createCell(1).setCellValue(shippingMethod);
            row.createCell(2).setCellValue(shippingCompany);
            row.createCell(3).setCellValue(waybillNumber.get(i));
            row.createCell(4).setCellValue(recieverName.get(i));
        }
    }

    private String extractExt(String originalFilename) {
        int pos = originalFilename.lastIndexOf(".");
        return originalFilename.substring(pos);
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