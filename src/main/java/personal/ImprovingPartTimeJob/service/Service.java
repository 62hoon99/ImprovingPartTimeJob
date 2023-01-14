package personal.ImprovingPartTimeJob.service;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.web.multipart.MultipartFile;
import personal.ImprovingPartTimeJob.repository.Repository;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@org.springframework.stereotype.Service
@RequiredArgsConstructor
public class Service {
    private final Repository repository;

    public String modifyOrderFile(MultipartFile file, String password) throws IOException {
        String storedFileName = repository.saveFile(file);
        repository.saveModifiedOrderFile(storedFileName, password);
        return storedFileName;
    }

    public String getFileDir() {
        return repository.getFileDir();
    }

    public void deleteFile(String storedFileName) {
        repository.deleteFile(storedFileName);
    }

    public String createBatchFile(MultipartFile orderFile, MultipartFile receiptFile) throws IOException {
        String orderFileName = repository.saveFile(orderFile);
        String receiptFileName = repository.saveFile(receiptFile);

        Workbook wbBatchFile = WorkbookFactory.create(new File(repository.getFixFileDir() + repository.getUploadFileName()));
        Workbook wbOrderFile = WorkbookFactory.create(new File(repository.getFileDir() + orderFileName));
        Workbook wbReceiptFile = WorkbookFactory.create(new File(repository.getFileDir() + receiptFileName));

        Sheet sheetBatch = wbBatchFile.getSheetAt(0);
        Sheet sheetOrder = wbOrderFile.getSheetAt(0);
        Sheet sheetReceipt = wbReceiptFile.getSheetAt(0);

        List<String> orderNumber = getColumnValues(sheetOrder, 0);
        List<String> waybillNumber = getColumnValues(sheetReceipt, 7);// H열
        List<String> payeeName = getColumnValues(sheetReceipt, 12); // M열

        setBatchFileValues(sheetBatch, orderNumber, waybillNumber, payeeName);

        String storedBatchFileName = getRandomStoredFileName();
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
            row.createCell(1).setCellValue(repository.getShippingMethod());
            row.createCell(2).setCellValue(repository.getShippingCompany());
            row.createCell(3).setCellValue(waybillNumber.get(i));
            row.createCell(4).setCellValue(recieverName.get(i));
        }
    }

    private String getRandomStoredFileName() {
        return UUID.randomUUID() + ".xls";
    }

    public void cleanUpDirectory() {
        repository.cleanupDirectory();
    }
}
