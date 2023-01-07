package personal.ImprovingPartTimeJob.service;

import lombok.RequiredArgsConstructor;
import org.springframework.web.multipart.MultipartFile;
import personal.ImprovingPartTimeJob.repository.Repository;

import java.io.IOException;

@org.springframework.stereotype.Service
@RequiredArgsConstructor
public class Service {
    private final Repository repository;

    public String modifyOrderFile(MultipartFile file) throws IOException {
        return repository.modifyOrderFile(file);
    }

    public String getFileDir() {
        return repository.getFileDir();
    }

    public void deleteFile(String storedFileName) {
        repository.deleteFile(storedFileName);
    }

    public String createBatchFile(MultipartFile orderFile, MultipartFile receiptFile, String turn) throws IOException {
        return repository.createBatchFile(orderFile, receiptFile, turn);
    }
}
