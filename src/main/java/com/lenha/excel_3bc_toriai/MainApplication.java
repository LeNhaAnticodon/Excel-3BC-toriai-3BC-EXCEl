package com.lenha.excel_3bc_toriai;

import com.lenha.excel_3bc_toriai.dao.SetupData;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;

import java.io.IOException;
import java.util.Objects;

public class MainApplication extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        // Tạo đối tượng Image từ file ảnh (đảm bảo file ảnh nằm trong thư mục resources)
        Image icon = new Image(Objects.requireNonNull(getClass().getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/LOGO_excel-3bc.png")));
        // Thiết lập biểu tượng cho Stage
        stage.getIcons().add(icon);

        FXMLLoader fxmlLoader = new FXMLLoader(MainApplication.class.getResource("convertExcelAnd3bc.fxml"));

        Scene scene = new Scene(fxmlLoader.load());
        stage.setTitle("CHUYỂN ĐỔI QUA LẠI GIỮA CÁC FILE EXCEL VÀ 3BC");
        stage.setScene(scene);
        stage.show();
        // lấy controller của FXMLLoader và gọi hàm getControls rồi thêm chính stage này vào list để hàm khởi tạo của
        // controller gọi hàm set language cho các control sẽ set ngôn ngữ cho chính title của stage này
        ((ConVertExcelAnd3BCController) fxmlLoader.getController()).getControls().add(stage);
    }

    @Override
    public void init() throws Exception {
        super.init();
        try {
            // đọc dữ liệu cài đặt từ file
            SetupData.getInstance().loadSetup();
        } catch (IOException e) {
            System.out.println("không đọc được file");
        }
    }

    public static void main(String[] args) {
        launch();
    }
}