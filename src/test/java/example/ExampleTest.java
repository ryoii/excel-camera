package example;

import cn.ryoii.ExcelCamera;
import cn.ryoii.ExcelCameraConfiguration;
import org.junit.Test;

import java.io.File;

public class ExampleTest {

    @Test
    public void example() {
        File file = new File("excel/example.xlsx");
        ExcelCameraConfiguration configuration = new ExcelCameraConfiguration(file);
        configuration
                .rowTo(12).colTo(13)
                .rowZoom(1.1).colZoom(1.1)
                .sheetName("example sheet");
        ExcelCamera excelCamera = new ExcelCamera(configuration);

        try {
            excelCamera.asImageFile("pic/example.png");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
