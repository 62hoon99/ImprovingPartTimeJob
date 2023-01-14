package personal.ImprovingPartTimeJob.controller.form;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

@Data
public class OrderFileForm {
    private MultipartFile orderFile;
    private String password;

    public OrderFileForm(String password) {
        this.password = password;
    }
}
