package com.lenha.excel_3bc_toriai;

import com.lenha.excel_3bc_toriai.convert.ExcelTo3BC;
import com.lenha.excel_3bc_toriai.convert.ReadPDFToExcel;
import com.lenha.excel_3bc_toriai.dao.SetupData;
import com.lenha.excel_3bc_toriai.model.ExcelFile;
import javafx.animation.KeyFrame;
import javafx.animation.Timeline;
import javafx.application.Platform;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.concurrent.Task;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.control.*;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.input.Clipboard;
import javafx.scene.input.ClipboardContent;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.paint.Color;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.util.Callback;
import javafx.util.Duration;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.text.NumberFormat;
import java.util.Locale;
import java.util.Map;
import java.util.Optional;
import java.util.ResourceBundle;
import java.util.concurrent.TimeoutException;

public class ConVertExcelAnd3BCController implements Initializable {

    @FXML
    public TextField linkExcelFile;
    @FXML
    public Button setExcelFileBtn;
    @FXML
    public TextField link3bcDir;
    @FXML
    public Button setSave3bcFileDirBtn;
    @FXML
    public ListView<ExcelFile> csvFIleList;
    @FXML
    public Button convertFileBtn;
    @FXML
    public Button openDirCsvBtn;
    @FXML
    public RadioButton setLangNihongoBtn;
    @FXML
    public ToggleGroup languages;
    @FXML
    public RadioButton setLangVietNamBtn;
    @FXML
    public RadioButton setLangEnglishBtn;
    @FXML
    public Menu menuHelp;
    @FXML
    public Menu menuEdit;
    @FXML
    public Menu menuFile;
    @FXML
    public Label listCsvFileTitle;
    @FXML
    public MenuBar menuBar;
    @FXML
    public Label copyLinkStatusLabel;
    @FXML
    public Button copyLinkBtn;
    @FXML
    public Label fileName;
    @FXML
    public Label product;
    @FXML
    public Label baseMaterial;
    @FXML
    public Button convertExcelTo3BCFileBtn;
    @FXML
    public TextField link3bcToriaiFile;
    @FXML
    public Button set3bcToriaiFileBtn;
    @FXML
    public TextField linkExcelResultDir;
    @FXML
    public Button setLinkExcelResultDirBtn;

    // map các ngôn ngữ
    private Map<String, String> languageMap;

    // bundle để lấy các giá trị trong file languagesMap. properties
    private ResourceBundle bundle;

    // các control cần thay đổi ngôn ngữ
    private ObservableList<Object> controls;

    // get các control cần thay đổi ngôn ngữ
    public ObservableList<Object> getControls() {
        return controls;
    }

    // alert của app
    private final Alert confirmAlert = new Alert(Alert.AlertType.CONFIRMATION);

    private static final String CONFIRM_PDF_FILE_LINK_TITLE = "Xác nhận địa chỉ file PDF";
    private static final String CONFIRM_PDF_FILE_LINK_HEADER = "Địa chỉ của file PDF chưa được xác nhận";
    private static final String CONFIRM_PDF_FILE_LINK_CONTENT = "Hãy chọn file PDF để tiếp tục!";

    private static final String CONFIRM_EXCEL_FILE_LINK_TITLE = "Xác nhận địa chỉ file Excel";
    private static final String CONFIRM_EXCEL_FILE_LINK_HEADER = "Địa chỉ của file Excel chưa được xác nhận";
    private static final String CONFIRM_EXCEL_FILE_LINK_CONTENT = "Hãy chọn file Excel để tiếp tục!";

    private static final String CONFIRM_CHL_FILE_DIR_TITLE = "Xác nhận thư mục chứa các file CHL";
    private static final String CONFIRM_CHL_FILE_DIR_HEADER = "Địa chỉ thư mục chứa các file CHL chưa được xác nhận";
    private static final String CONFIRM_CHL_FILE_DIR_CONTENT = "Hãy chọn thư mục chứa để tiếp tục!";

    private static final String CONFIRM_3BC_FILE_DIR_TITLE = "Xác nhận thư mục chứa các file 3BC";
    private static final String CONFIRM_3BC_FILE_DIR_HEADER = "Địa chỉ thư mục chứa các file 3BC chưa được xác nhận";
    private static final String CONFIRM_3BC_FILE_DIR_CONTENT = "Hãy chọn thư mục chứa để tiếp tục!";

    private static final String CONFIRM_CONVERT_COMPLETE_TITLE = "Thông tin hoạt động chuyển file";
    private static final String CONFIRM_CONVERT_COMPLETE_HEADER = "Đã chuyển xong file PDF sang các file CHL";
    // NEW
    private static final String CONFIRM_CONVERT_EXCEL_TO_3BC_COMPLETE_HEADER = "Đã chuyển xong file EXCEL sang file 3BC";
    private static final String CO_VAT_LIEU_KHONG_TON_TAI = "Có tồn tại vật liệu không nằm trong danh sách";
    private static final String CONFIRM_CONVERT_COMPLETE_CONTENT = "Bạn có muốn mở thư mục chứa các file CHL và\ntự động copy địa chỉ không?";
    // NEW
    private static final String CONFIRM_CONVERT_EXCEL_TO_3BC_COMPLETE_CONTENT = "Bạn có muốn mở thư mục chứa file 3BC và\ntự động copy địa chỉ không?";
    private static final String THONG_BAO_DOI_VAT_LIEU = "Đã thay thế tất cả vật liệu sang bộ vật liệu dự phòng!";

    private static final String ERROR_CONVERT_TITLE = "Thông báo lỗi chuyển file";
    private static final String ERROR_CONVERT_HEADER = "Nội dung file PDF không phải là tính toán vật liệu hoặc file không được phép truy cập";
    //NEW
    private static final String ERROR_CONVERT_EXCEL_TO_3BC_HEADER = "Nội dung file EXCEL không phải là tính toán vật liệu hoặc file không được phép truy cập";
    private static final String ERROR_CONVERT_CONTENT = "Bạn có muốn chọn file khác và thực hiện lại không?";

    private static final String ERROR_OPEN_CHL_DIR_TITLE = "Lỗi mở thư mục";
    private static final String ERROR_COPY_CHL_DIR_TITLE = "Lỗi copy địa chỉ thư mục";
    private static final String ERROR_CHL_DIR_HEADER = "Thư mục chứa các file CHL có địa chỉ không đúng hoặc chưa được chọn!";
    private static final String ERROR_COPY_CHL_DIR_CONTENT = "Không thể copy địa chỉ thư mục chứa các file CHL";


    /**
     * hàm khởi tạo
     *
     * @param url            The location used to resolve relative paths for the root object, or
     *                       {@code null} if the location is not known.
     * @param resourceBundle The resources used to localize the root object, or {@code null} if
     *                       the root object was not localized.
     */
    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {

        System.out.println("link excel " + SetupData.getInstance().getSetup().getLinkExcelFile());
        System.out.println("link toriai 3bc " + SetupData.getInstance().getSetup().getLink3bcToriaiFile());
        System.out.println("link 3bc dir " + SetupData.getInstance().getSetup().getLinkSave3BCFileDir());

        // bind list view với list các file chl đã tạo
        csvFIleList.setItems(SetupData.getInstance().getExcelFiles());
        // cài đặt các cell của list view
        setupCellChlFIleList();

        // lấy map chứa các câu bằng các ngôn ngữ khác nhau và value là từ khóa của chúng trong file languagesMap.properties
        languageMap = SetupData.getInstance().getLanguageMap();
        // lấy list chứa các control UI cần để thay đổi ngôn ngữ hiển thị
        controls = SetupData.getInstance().getControls();
        // thêm các control của controller này vào map
        controls.addAll(setExcelFileBtn, setSave3bcFileDirBtn, set3bcToriaiFileBtn, setLinkExcelResultDirBtn, convertFileBtn, openDirCsvBtn, listCsvFileTitle, menuBar, copyLinkStatusLabel, copyLinkBtn,
                fileName, product, baseMaterial);

        // Lấy bundle của file ngôn ngữ
        bundle = ResourceBundle.getBundle("languagesMap");

        // set UserData cho 3 radio button ngôn ngữ
        setLangVietNamBtn.setUserData("vi");
        setLangEnglishBtn.setUserData("en");
        setLangNihongoBtn.setUserData("ja");

        // lấy ngôn ngữ đã lưu trong file setup
        String lang = SetupData.getInstance().getSetup().getLang();
        // nếu lang không có giá trị hoặc giá trị không đúng cấu trúc thì cho mặc định là tiếng Việt
        // và cho radio click vào nút tiếng việt
        if (lang.isBlank() || (!lang.equals("vi") && !lang.equals("en") && !lang.equals("ja"))) {
            languages.selectToggle(setLangVietNamBtn);
        }
        // còn không thì cho radio click vào nút tương ứng với lang
        else {
            if (lang.equals("vi")) {
                languages.selectToggle(setLangVietNamBtn);
            }

            if (lang.equals("en")) {
                languages.selectToggle(setLangEnglishBtn);
            }

            if (lang.equals("ja")) {
                languages.selectToggle(setLangNihongoBtn);
            }
        }

        // cho hiển thị ngôn ngữ các control theo radio button ngôn ngữ đã chọn
        updateLangInBackground(languages.getSelectedToggle(), controls);
//        setlang();

        // tạo sự kiện thay đổi nút ngôn ngữ thì cho hiển thị theo ngôn ngữ mới
        // và lưu ngôn ngữ mới vào file và đối tượng setup
        languages.selectedToggleProperty().addListener((observableValue, oldValue, newValue) -> {
            if (newValue != null) {
                updateLangInBackground(newValue, controls);
                // và lưu ngôn ngữ mới vào file và đối tượng setup
                setlang();
            }
        });

        // nếu 3 ô hiển thị link file, thư mục có thay đổi giá trị thì cho viền của chúng đổi màu trong 3s
        linkExcelFile.textProperty().addListener((observableValue, oldValue, newValue) -> {
            setBorderColorTF(linkExcelFile);
        });
        link3bcToriaiFile.textProperty().addListener((observableValue, oldValue, newValue) -> {
            setBorderColorTF(link3bcToriaiFile);
        });
        link3bcDir.textProperty().addListener((observableValue, oldValue, newValue) -> {
            setBorderColorTF(link3bcDir);
        });

        // lấy địa chỉ đã chọn lần gần nhất của thư mục sẽ lưu file 3bc
        File file3bcDiv = new File(SetupData.getInstance().getSetup().getLinkSave3BCFileDir());

        // lấy địa chỉ đã chọn lần gần nhất của thư mục sẽ lưu file excel
        File fileExcelDiv = new File(SetupData.getInstance().getSetup().getLinkSaveExcelFileDir());

        // nếu địa chỉ đúng là thư mục thì cho hiển thị trên màn hình
        if (fileExcelDiv.isDirectory()) {
            linkExcelResultDir.setText(SetupData.getInstance().getSetup().getLinkSaveExcelFileDir());
        }

        // nếu địa chỉ đúng là thư mục thì cho hiển thị trên màn hình
        if (file3bcDiv.isDirectory()) {
            link3bcDir.setText(SetupData.getInstance().getSetup().getLinkSave3BCFileDir());
        }
    }

    /**
     * đổi màu viền của textfield hoặc text area
     */
    private void setBorderColorTF(TextInputControl textField) {
        // chạy trong nền
        Task<Void> task = new Task<>() {
            @Override
            protected Void call() {
                Platform.runLater(() -> {
                    // Đổi màu viền
                    textField.setStyle("-fx-border-color: #FFA000; -fx-border-width: 2; -fx-border-radius: 5");
                    // Tạo Timeline để xóa viền sau 3 giây
                    Timeline timeline = new Timeline(new KeyFrame(
                            Duration.seconds(3),
                            event -> {
                                textField.setStyle("-fx-border-color:  none");
//                                textField.setStyle("-fx-border-width:  none");
//                                textField.setStyle("-fx-border-radius:  none");
                            }
                    ));
                    // Chạy Timeline một lần
                    timeline.setCycleCount(1);
                    timeline.play();

                });
                return null;
            }
        };

        Thread thread = new Thread(task);
        thread.setDaemon(true);
        thread.start();
    }

    /**
     * set ngôn ngữ theo nút ngôn ngữ đang chọn cho đối tượng cài đặt và lưu ngôn ngữ vào file
     */
    private void setlang() {
        try {
            SetupData.getInstance().setLang(languages.getSelectedToggle().getUserData().toString());
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * lấy địa chỉ file excel cần tính vật liệu
     *
     * @return file pdf đã chọn
     */
    @FXML
    public File getExcelFile() {
        // tạo trình chọn file excel
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("EXCEL", "*.xlsx"));

        // lấy địa chỉ file excel đã chọn lần trước
        String oldLinkExcelFile = SetupData.getInstance().getSetup().getLinkExcelFile();
        System.out.println("old" + oldLinkExcelFile);
        // nếu địa chỉ không rỗng thì lấy thư mục chứa file lần trước và cho trình chọn file mở thư mục đó
        if (!oldLinkExcelFile.isBlank()) {
            // tách link tại các đoạn phân tách địa chỉ
            String[] oldLinkExcelFileArr = oldLinkExcelFile.split("\\\\");
            // link thư mục chứa file pdf
            String linkExcelFileDir = "";
            // ghép lại link nhưng bỏ phần cuối cùng(tên file xlsx) là ra địa chỉ thư mục chứa file pdf
            for (int i = 0; i < oldLinkExcelFileArr.length - 1; i++) {
                linkExcelFileDir = linkExcelFileDir.concat(oldLinkExcelFileArr[i]).concat("\\");
            }

            // tạo file thư mục chứa file excel
            File file = new File(linkExcelFileDir);
            // nếu file là thư mục thì cho trình chọn file pdf bắt đầu chọn file từ thư mục này
            if (file.isDirectory()) {
                fileChooser.setInitialDirectory(file);
            }
        }

        // file excel
        File file;
        // chọn file pdf nếu không có lỗi địa chỉ khởi đầu thì thoát,nếu không cho địa chỉ khởi đầu là null và chọn lại đến khi hết lỗi
        while (true) {
            try {
                // lấy file excel
                file = fileChooser.showOpenDialog(menuBar.getScene().getWindow());
                break;
            } catch (IllegalArgumentException e) {
                System.out.println("địa chỉ khởi đầu không hợp lệ");
                fileChooser.setInitialDirectory(null);
            }
        }

        // nếu file pdf không null và là file hợp lệ
        if (file != null && file.isFile()) {
            // địa chỉ file
            String link = file.getAbsolutePath();
            // nếu địa chỉ link của file được chọn khác với địa chỉ cũ đang được chọn thì xóa danh sách list các file chl đã được tạo(file cũ là excel)
            if (!link.equals(SetupData.getInstance().getSetup().getLinkExcelFile())) {
                SetupData.getInstance().getExcelFiles().clear();
            }
            // cho hiển thị địa chỉ file rồi lưu vào đối tượng cài đặt và lưu vào file
            linkExcelFile.setText(link);
            SetupData.getInstance().setLinkExcelFile(link);

//            // nếu địa chỉ thư mục chứa file excel sẽ tạo chưa được chọn thì chọn luôn thư mục của file excel vừa chọn
//            // không đúng logic của app này
//            if (link3bcDir.getText().isBlank()) {
//                // phân tách địa chỉ file pdf
//                String[] csvDirArr = link.split("\\\\");
//                // địa chỉ thư mục chứa file chl
//                String csvDir = "";
//                // ghép lại link nhưng bỏ phần cuối cùng(tên file pdf) là ra địa chỉ thư mục chứa file pdf
//                for (int i = 0; i < csvDirArr.length - 1; i++) {
//                    csvDir = csvDir.concat(csvDirArr[i]).concat("\\");
//                }
//                // hiển thị địa chỉ thư mục
//                link3bcDir.setText(csvDir);
//                // lưu địa chỉ vào đối tượng setup và lưu vào file
//                SetupData.getInstance().setLinkSave3bcFileDir(csvDir);
//            }
        } else {
            System.out.println("không chọn file");
        }

        return file;
    }

    /**
     * lấy địa chỉ file pdf tính vật liệu của 3bc
     *
     * @return
     */
    @FXML
    public File get3bcToriaiFile() {
        // tạo trình chọn file pdf tính vật liệu
        FileChooser fileChooser = new FileChooser();
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("PDF", "*.pdf"));

        // lấy địa chỉ file pdf đã chọn lần trước
        String oldLinkPdfFile = SetupData.getInstance().getSetup().getLink3bcToriaiFile();
        System.out.println("old" + oldLinkPdfFile);
        // nếu địa chỉ không rỗng thì lấy thư mục chứa file lần trước và cho trình chọn file mở thư mục đó
        if (!oldLinkPdfFile.isBlank()) {
            // tách link tại các đoạn phân tách địa chỉ
            String[] oldLinkPdfFileArr = oldLinkPdfFile.split("\\\\");
            // link thư mục chứa file pdf
            String linkPdfFileDir = "";
            // ghép lại link nhưng bỏ phần cuối cùng(tên file pdf) là ra địa chỉ thư mục chứa file pdf
            for (int i = 0; i < oldLinkPdfFileArr.length - 1; i++) {
                linkPdfFileDir = linkPdfFileDir.concat(oldLinkPdfFileArr[i]).concat("\\");
            }

            // tạo file thư mục chứa file pdf
            File file = new File(linkPdfFileDir);
            // nếu file là thư mục thì cho trình chọn file pdf bắt đầu chọn file từ thư mục này
            if (file.isDirectory()) {
                fileChooser.setInitialDirectory(file);
            }
        }

        // file pdf
        File file;
        // chọn file pdf nếu không có lỗi địa chỉ khởi đầu thì thoát,nếu không cho địa chỉ khơ đầu là null và chọn lại đến khi hết lỗi
        while (true) {
            try {
                // lấy file pdf
                file = fileChooser.showOpenDialog(menuBar.getScene().getWindow());
                break;
            } catch (IllegalArgumentException e) {
                System.out.println("địa chỉ khởi đầu không hợp lệ");
                fileChooser.setInitialDirectory(null);
            }
        }

        // nếu file pdf không null và là file hợp lệ
        if (file != null && file.isFile()) {
            // địa chỉ file
            String link = file.getAbsolutePath();

            // cho hiển thị địa chỉ file rồi lưu vào đối tượng cài đặt và lưu vào file
            link3bcToriaiFile.setText(link);
            SetupData.getInstance().setLink3bcToriaiFile(link);

        } else {
            System.out.println("không chọn file");
        }

        return file;
    }

    /**
     * lấy địa chỉ thư mục chứa file 3bc
     *
     * @return file thư mục chứa file 3bc đã chọn
     */
    @FXML
    public File setSave3bcFileDir() {
        // trình chọn thư mục
        DirectoryChooser directoryChooser = new DirectoryChooser();
        // thư mục chứa file 3bc lần trước
        String oldDir = SetupData.getInstance().getSetup().getLinkSave3BCFileDir();
        File oldFileDir = new File(oldDir);
        // nếu thư mục chứa file 3bc lần trước là thư mục thì cho trình chọn thư mục bắt đầu từ thư mục này
        if (oldFileDir.isDirectory()) {
            directoryChooser.setInitialDirectory(oldFileDir);
        }

        // lấy thư mục chứa file 3bc vừa chọn
        File dir = directoryChooser.showDialog(menuBar.getScene().getWindow());

        // nếu thư mục chứa file 3bc không null và là thư mục hợp lệ thì hiển thị rồi lưu địa chỉ vào đối tượng setup và lưu vào file
        if (dir != null && dir.isDirectory()) {

            String link = dir.getAbsolutePath();

//            // nếu địa chỉ link của thư mục được chọn khác với địa chỉ cũ đang được chọn thì xóa danh sách list các file excel
//            // không đúng logic với ứng dụng này
//            if (!link.equals(SetupData.getInstance().getSetup().getLinkSave3bcFileDir())) {
//                SetupData.getInstance().getChlFiles().clear();
//            }

            // hiển thị link
            link3bcDir.setText(link);
            // lưu vào đối tượng setup và lưu vào file
            SetupData.getInstance().setLinkSave3bcFileDir(link);
        } else {
            System.out.println("không chọn thư mục");
        }
        // trả về thư mục vừa chọn
        return dir;
    }


    public File setSaveExcelFileDir(ActionEvent actionEvent) {
        // trình chọn thư mục
        DirectoryChooser directoryChooser = new DirectoryChooser();
        // thư mục chứa file 3bc lần trước
        String oldDir = SetupData.getInstance().getSetup().getLinkSaveExcelFileDir();
        File oldFileDir = new File(oldDir);
        // nếu thư mục chứa file 3bc lần trước là thư mục thì cho trình chọn thư mục bắt đầu từ thư mục này
        if (oldFileDir.isDirectory()) {
            directoryChooser.setInitialDirectory(oldFileDir);
        }

        // lấy thư mục chứa file 3bc vừa chọn
        File dir = directoryChooser.showDialog(menuBar.getScene().getWindow());

        // nếu thư mục chứa file 3bc không null và là thư mục hợp lệ thì hiển thị rồi lưu địa chỉ vào đối tượng setup và lưu vào file
        if (dir != null && dir.isDirectory()) {

            String link = dir.getAbsolutePath();

//            // nếu địa chỉ link của thư mục được chọn khác với địa chỉ cũ đang được chọn thì xóa danh sách list các file excel
//            // không đúng logic với ứng dụng này
//            if (!link.equals(SetupData.getInstance().getSetup().getLinkSave3bcFileDir())) {
//                SetupData.getInstance().getChlFiles().clear();
//            }

            // hiển thị link
            linkExcelResultDir.setText(link);
            // lưu vào đối tượng setup và lưu vào file
            SetupData.getInstance().setLinkSaveExcelFileDir(link);
        } else {
            System.out.println("không chọn thư mục");
        }
        // trả về thư mục vừa chọn
        return dir;
    }

    /**
     * thực hiện chuyển dữ liệu từ file pdf sang file excel
     */
    @FXML
    public void convertFilePdfToriaiToExcel(ActionEvent actionEvent) {
        // link file pdf
        String pdfFilePath;
        // link thư mục chứa các file chl
        String chlFileDirPath;

        // yêu cầu chọn địa chỉ file và thư mục khi 2 địa chỉ này chưa được chọn
        // nếu chọn xong thì phải chuyển dữ liệu thành công thì mới thoát được vòng lặp
        while (true) {
            // lấy link từ các ô đang hiển thị
            pdfFilePath = linkExcelFile.getText();
            chlFileDirPath = link3bcDir.getText();

            // tạo file theo các link trên
            File pdfFile = new File(pdfFilePath);
            File chlFileDir = new File(chlFileDirPath);

            // kiểm tra hợp lệ địa chỉ file và thư mục
            boolean isFilePDF = pdfFile.isFile();
            boolean isDir = chlFileDir.isDirectory();

            // nếu không phải là file pdf thì yêu cầu chọn lại
            if (!isFilePDF) {
                // hiển thị alert yêu cầu chọn lại file pdf
                confirmAlert.setTitle(CONFIRM_PDF_FILE_LINK_TITLE);
                confirmAlert.setHeaderText(CONFIRM_PDF_FILE_LINK_HEADER);
                confirmAlert.setContentText(CONFIRM_PDF_FILE_LINK_CONTENT);
                updateLangAlert(confirmAlert);
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu đồng ý thì gọi hàm chọn file
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File fileSelected = getExcelFile();
                    // nếu file đã chọn null nghĩa là người dùng click vào nút cancel khi chọn file
                    // thì thoát và không làm gì nữa
                    if (fileSelected == null) {
                        return;
                    }

                    /* sau khi đã chọn xong file pdf thì gán luôn file do hàm chọn trả về để hàm lấy file cho việc chuyển bên dưới
                     hoạt động đúng mà không phải thêm 1 vòng lặp nữa tính lại file */
                    pdfFile = fileSelected;

                    // nếu file chọn không đúng thì nhảy sang sang vòng lặp mới và chọn lại từ đầu
                    if (!pdfFile.isFile()) {
                        continue;
                    }
                }
                // nếu không đồng ý thì thoát hàm
                else {
                    return;
                }
            }

            // nếu không phải là thư mục thì yêu cầu chọn lại
            if (!isDir) {
                /* nếu địa chỉ file pdf đã xác nhận thì nó sẽ tự động lấy địa chỉ thư mục chứa file pdf đó nhập vào
                 linkCvsDir, mà trước đó đã xác nhận chưa chọn thư mục chứa file đã chuyển nên cần xóa text của
                 linkCvsDir đi để người dùng xác nhận lại, tránh hiển thị địa chỉ mặc định trên ỏ linkCvsDir làm khó hiểu */
                link3bcDir.setText("");
                SetupData.getInstance().setLinkSave3bcFileDir("");
                // hiển thị alert yêu cầu chọn lại
                confirmAlert.setTitle(CONFIRM_CHL_FILE_DIR_TITLE);
                confirmAlert.setHeaderText(CONFIRM_CHL_FILE_DIR_HEADER);
                confirmAlert.setContentText(CONFIRM_CHL_FILE_DIR_CONTENT);
                updateLangAlert(confirmAlert);
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu là nút ok thì gọi hàm chọn thư mục
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File dirSelected = setSave3bcFileDir();

                    // nếu thư mục trả về null tức người dùng hủy chọn bằng nút cancel thì thoát hàm và không làm gì nữa
                    if (dirSelected == null) {
                        return;
                    }

                    /* sau khi đã chọn xong thư mục thì gán luôn thư mục do hàm chọn trả về để hàm lấy thư mục cho việc chuyển bên dưới
                     hoạt động đúng mà không phải thêm 1 vòng lặp nữa tính lại thư mục */
                    chlFileDir = dirSelected;

                    // nếu thư mục trả về không đúng thì nhảy sang sang vòng lặp mới và chọn lại từ đầu
                    if (!chlFileDir.isDirectory()) {
                        continue;
                    }
                }
                // nếu không đồng ý thì thoát hàm
                else {
                    return;
                }
            }

            // đến đây nếu không bị return thì đã chọn xong 2 địa chỉ file
            if (pdfFile.isFile() && chlFileDir.isDirectory()) {
                System.out.println("đã chọn xong 2 địa chỉ");
                System.out.println(pdfFile.getAbsolutePath());
                System.out.println(chlFileDir.getAbsolutePath());
            }

            // gọi hàm chuyển file từ class static ReadPDFToExcel
            try {
                ReadPDFToExcel.convertPDFToExcel(pdfFile.getAbsolutePath(), chlFileDir.getAbsolutePath(), SetupData.getInstance().getExcelFiles());
                // hiển thị alert chuyển file thành công
                confirmAlert.setTitle(CONFIRM_CONVERT_COMPLETE_TITLE);
                confirmAlert.setHeaderText(CONFIRM_CONVERT_COMPLETE_HEADER);
                confirmAlert.setContentText(CONFIRM_CONVERT_COMPLETE_CONTENT);

                updateLangAlert(confirmAlert);

                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu là nút ok thì copy đường dẫn thư mục chứa file chl vào clipboard và mở thư mục này
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    copyContentToClipBoard(chlFileDir.getAbsolutePath());
                    // mở thư mục chứa file chl
                    Desktop.getDesktop().open(chlFileDir);
                }

                return;
            } catch (Exception e) {
                System.out.println(e.getMessage());
                e.printStackTrace();
                // xóa hết các button, đổi alert sang dạng error rồi thêm lại 2 nút ok và cancel
                confirmAlert.getButtonTypes().clear();
                confirmAlert.setAlertType(Alert.AlertType.ERROR);
                confirmAlert.getButtonTypes().add(ButtonType.CANCEL);
                confirmAlert.getButtonTypes().add(ButtonType.OK);

                confirmAlert.setTitle(ERROR_CONVERT_TITLE);
                confirmAlert.setHeaderText(ERROR_CONVERT_HEADER);
                confirmAlert.setContentText(ERROR_CONVERT_CONTENT);
                // cập nhật ngôn ngữ cho alert
                updateLangAlert(confirmAlert);

                // nếu là sự kiện không ghi được file chl do file trùng tên với file sắp tạo đang được mở
                // th in ra cảnh báo và thoát
                if (e instanceof FileNotFoundException) {
                    confirmAlert.getButtonTypes().clear();
                    confirmAlert.getButtonTypes().add(ButtonType.OK);
                    confirmAlert.setHeaderText("Tên file CHL đang tạo: (\"" + ReadPDFToExcel.fileName + "\") trùng tên với 1 file CHL khác đang được mở nên không thể ghi đè");
                    confirmAlert.setContentText("Hãy đóng file CHL đang mở để tiếp tục!");
                    System.out.println("File đang được mở bởi người dùng khác");
                    updateLangAlert(confirmAlert);
                    confirmAlert.showAndWait();

                    // chuyển lại alert về dạng confirm và thêm nút cancel
                    confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                    confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

                    return;
                }
                // nếu là lỗi quá 99 dòng thì thông báo và thoát
                if (e instanceof TimeoutException) {
                    confirmAlert.getButtonTypes().clear();
                    confirmAlert.getButtonTypes().add(ButtonType.OK);
                    confirmAlert.setHeaderText("File CHL đang tạo: (\"" + ReadPDFToExcel.fileName + "\") trong một boZai duy nhất có số dòng sản phẩm cần ghi lớn hơn 99 nên không thể ghi");
                    confirmAlert.setContentText("Hãy chỉnh sửa lại dữ liệu vật liệu đang chuyển để tiếp tục!");
                    System.out.println("Vật liệu có số dòng lớn hơn 99");
                    updateLangAlert(confirmAlert);
                    confirmAlert.showAndWait();

                    // chuyển lại alert về dạng confirm và thêm nút cancel
                    confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                    confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

                    return;
                }


                // nếu là lỗi ghi file thì thông báo
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // chuyển lại alert về dạng confirm
                confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);

                // nếu chọn ok thì gọi lại hàm chọn file pdf để chọn file khác
                // nếu chọn cancel thì thoát
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File fileSelected2 = getExcelFile();

                    // nếu không chọn file thì thoát
                    if (fileSelected2 == null) {
                        return;
                    }
                } else {
                    return;
                }
            }
        }

    }

    /**
     * thay đổi ngôn ngữ của alert theo ngôn ngữ đang chọn
     *
     * @param alert cần thay đổi ngôn ngữ
     */
    private void updateLangAlert(Alert alert) {
        // gọi hàm update ngôn ngữ trong nền
        updateLangInBackground(languages.getSelectedToggle(), FXCollections.observableArrayList(alert));
    }

    /**
     * mở cửa sổ chọn thư mục chứa file excel đang được hiển thị khi chuyển file xong
     */
    @FXML
    public void openExcelDir(ActionEvent actionEvent) {
        // lấy địa chỉ thư mục đang hiển thị gán vào file
        File excelFileDir = new File(linkExcelResultDir.getText());
        // nếu file là thư mục thì mở thư mục bằng cửa sổ của window
        // nếu không thì thông báo lỗi
        if (excelFileDir.isDirectory()) {
            try {
                Desktop.getDesktop().open(excelFileDir);
            } catch (IOException e) {
                System.out.println(e.getMessage());
                confirmAlert.setAlertType(Alert.AlertType.ERROR);
                confirmAlert.setTitle(ERROR_OPEN_CHL_DIR_TITLE);
                confirmAlert.setHeaderText(e.getMessage());
                confirmAlert.setContentText("");
                updateLangAlert(confirmAlert);
                confirmAlert.showAndWait();
                confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

            }
        } else {
            confirmAlert.setAlertType(Alert.AlertType.ERROR);
            confirmAlert.setTitle(ERROR_OPEN_CHL_DIR_TITLE);
            confirmAlert.setHeaderText(ERROR_CHL_DIR_HEADER);
            confirmAlert.setContentText("");
            updateLangAlert(confirmAlert);
            confirmAlert.showAndWait();
            confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
            confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

        }

    }

    /**
     * mở cửa sổ chọn thư mục chứa file 3bc đang được hiển thị khi chuyển file xong
     */
    @FXML
    public void open3bcFileDir(ActionEvent actionEvent) {
        // lấy địa chỉ thư mục đang hiển thị gán vào file
        File _3bcFileDir = new File(link3bcDir.getText());
        // nếu file là thư mục thì mở thư mục bằng cửa sổ của window
        // nếu không thì thông báo lỗi
        if (_3bcFileDir.isDirectory()) {
            try {
                Desktop.getDesktop().open(_3bcFileDir);
            } catch (IOException e) {
                System.out.println(e.getMessage());
                confirmAlert.setAlertType(Alert.AlertType.ERROR);
                confirmAlert.setTitle(ERROR_OPEN_CHL_DIR_TITLE);
                confirmAlert.setHeaderText(e.getMessage());
                confirmAlert.setContentText("");
                updateLangAlert(confirmAlert);
                confirmAlert.showAndWait();
                confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

            }
        } else {
            confirmAlert.setAlertType(Alert.AlertType.ERROR);
            confirmAlert.setTitle(ERROR_OPEN_CHL_DIR_TITLE);
            confirmAlert.setHeaderText(ERROR_CHL_DIR_HEADER);
            confirmAlert.setContentText("");
            updateLangAlert(confirmAlert);
            confirmAlert.showAndWait();
            confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
            confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

        }

    }

    @FXML
    public void setLangNihongo(ActionEvent actionEvent) {
    }

    @FXML
    public void setLangVietNam(ActionEvent actionEvent) {
    }

    @FXML
    public void setLangEnglish(ActionEvent actionEvent) {
    }

    /**
     * cập nhật ngôn ngữ trong nền
     *
     * @param langBtn  nút radio của ngôn ngữ đang chọn
     * @param controls các control cần update ngôn ngữ
     */
    public void updateLangInBackground(Toggle langBtn, ObservableList<Object> controls) {
        // lấy user data của nút ngôn ngữ
        String lang = langBtn.getUserData().toString();
        // tạo tác vụ chạy nền và gọi hàm update ngôn ngữ
        Task<Void> task = new Task<>() {
            @Override
            protected Void call() {
                Platform.runLater(() -> updateLang(lang, controls));
                return null;
            }
        };

        Thread thread = new Thread(task);
        thread.setDaemon(true);
        thread.start();
    }

    /**
     * cập nhật text của các control theo ngôn ngữ đang chọn
     *
     * @param lang     ngôn ngữ đang chọn
     * @param controls các control cần update ngôn ngữ
     */
    private void updateLang(String lang, ObservableList<Object> controls) {
        // duyệt qua các control và thay đổi text của nó theo ngôn ngữ đang chọn
        for (Object control : controls) {
            // nếu control là label thì thay đổi ngôn ngữ bằng hàm setText
            // các control khác cần thay đổi các text khác nhau nhưng cách làm tương tự
            if (control instanceof Labeled labeledControl) {
                // lấy text đang hiển thị của label
                String currentText = labeledControl.getText();
                // lấy key của text này trong map ngôn ngữ
                String key = languageMap.get(currentText);
                // từ key này lấy ra từ tương ứng theo ngôn ngữ đang hiển thị trong file bundle languageMap.properties
                if (key != null) {
                    String newText = bundle.getString(key + "." + lang);
                    labeledControl.setText(newText);
                }
            } else if (control instanceof MenuBar menuBar1) {
                for (Menu menu : menuBar1.getMenus()) {
                    String currentText = menu.getText();
                    String key = languageMap.get(currentText);
                    if (key != null) {
                        String newText = bundle.getString(key + "." + lang);
                        menu.setText(newText);
                    }
                    for (MenuItem menuItem : menu.getItems()) {
                        currentText = menuItem.getText();
                        key = languageMap.get(currentText);
                        if (key != null) {
                            String newText = bundle.getString(key + "." + lang);
                            menuItem.setText(newText);
                        }
                    }
                }
            } else if (control instanceof Alert alert) {
                String title = alert.getTitle();
                String header = alert.getHeaderText();
                String content = alert.getContentText();

                String fileName = "";
                // nếu header có .sysc2 tức là trong tên có tên file đang tạo bị lỗi hoặc file có số dòng lớn hơn 99
                // ở sự kiện tên file sắp tạo trùng tên với file đang mở
                // tách tên file ra ghi vào fileName
                // chỉ lấy phần cố định thêm "" vào giữa gán cho header
                // phần cố định sẽ có trong map languageMap và lấy được keyHeader trong languageMap
                // từ keyHeader lấy được ngôn ngữ đang dùng trong bundle
                // phần tách tiếp ngôn ngữ chia 2 nửa tại điểm " rồi thêm " + fileName + " vào giữa để hiển thị hoàn chỉnh theo ngôn ngữ này
                if (header.contains(".sysc2") || header.contains(".csv") || header.contains(".xlsx")) {
                    String[] headerarr = header.split("\"");
                    fileName = headerarr[1];
                    header = headerarr[0] + "\"\"" + headerarr[2];
                }

                String keyTitle = languageMap.get(title);
                String keyHeader = languageMap.get(header);
                String keyContent = languageMap.get(content);


                if (keyTitle != null) {
                    alert.setTitle(bundle.getString(keyTitle + "." + lang));
                }

                if (keyHeader != null) {
                    if (fileName.isBlank()) {
                        alert.setHeaderText(bundle.getString(keyHeader + "." + lang));
                    } else {
                        String[] headerArr = bundle.getString(keyHeader + "." + lang).split("\"");
                        alert.setHeaderText(headerArr[0] + "\"" + fileName + "\"" + headerArr[2]);

                    }
                }

                if (keyContent != null) {
                    alert.setContentText(bundle.getString(keyContent + "." + lang));
                }
            } else if (control instanceof Stage stage) {
                String title = stage.getTitle();
                String keyTitle = languageMap.get(title);
                if (keyTitle != null) {
                    stage.setTitle(bundle.getString(keyTitle + "." + lang));
                }
            } else if (control instanceof Dialog dialog) {
                String title = dialog.getTitle();
                String keyTitle = languageMap.get(title);
                if (keyTitle != null) {
                    dialog.setTitle(bundle.getString(keyTitle + "." + lang));
                }
            }
        }
    }

    /**
     * copy nội dung content vào clipboard của window
     *
     * @param content nội dung cần copy
     */
    private void copyContentToClipBoard(String content) {
        Clipboard clipboard = Clipboard.getSystemClipboard();
        ClipboardContent clipboardContent = new ClipboardContent();
        clipboardContent.putString(content);
        clipboard.setContent(clipboardContent);
    }

    /**
     * copy địa chỉ thư mục chứa file excel kết quả
     */
    public void copyLinkExcelResultDir(ActionEvent actionEvent) {
        // tạo tác vụ chạy nền và gọi hàm copy địa chỉ thư mục chứa các file chl
        Task<Void> task = new Task<>() {
            @Override
            protected Void call() {
                Platform.runLater(() -> copylinkExcelResultFolder());
                return null;
            }
        };

        Thread thread = new Thread(task);
        thread.setDaemon(true);
        thread.start();

    }

    /**
     * copy địa chỉ thư mục chứa file excel kết quả
     */
    private void copylinkExcelResultFolder() {

        // lấy địa chỉ thư mục chứa file excel
        File chlFileDir = new File(link3bcDir.getText());
        // nếu thư mục chứa các file chl là thư mục thì copy địa chỉ thư mục chứa các file chl vào clipboard
        // hiển thị label thông báo đã copy trong 3 giây
        if (chlFileDir.isDirectory()) {
            // gọi hàm copy
            copyContentToClipBoard(chlFileDir.getAbsolutePath());

            // hiển thị label thông báo đã copy trong 3 giây
            copyLinkStatusLabel.setVisible(true);
            // Tạo Timeline để ẩn Label sau 3 giây
            Timeline timeline = new Timeline(new KeyFrame(
                    Duration.seconds(3),
                    event -> copyLinkStatusLabel.setVisible(false) // Ẩn Label sau 3 giây
            ));
            // Chạy Timeline một lần
            timeline.setCycleCount(1);
            timeline.play();
        } else {
            confirmAlert.setAlertType(Alert.AlertType.ERROR);
            confirmAlert.setTitle(ERROR_COPY_CHL_DIR_TITLE);
            confirmAlert.setHeaderText(ERROR_CHL_DIR_HEADER);
            confirmAlert.setContentText(ERROR_COPY_CHL_DIR_CONTENT);
            updateLangAlert(confirmAlert);
            confirmAlert.showAndWait();
            confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
            confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

        }

    }

    /**
     * cài đặt định dạng cho các cell của list view
     */
    private void setupCellChlFIleList() {
        /*gọi hàm setCellFactory để cài đặt lại các thuộc tính của ListView
            tham số là 1 FunctionalInterface Callback, ta sẽ tạo lớp ẩn danh của
            Interface này để Override method call của nó
            cần xác định các thuộc tính để Callback truyền vào cho method call bằng generics với
            2 thuộc tính lần này là ListView<ExcelFile> và ListCell<ExcelFile>*/
        csvFIleList.setCellFactory(new Callback<ListView<ExcelFile>, ListCell<ExcelFile>>() {
            @Override
            public ListCell<ExcelFile> call(ListView<ExcelFile> CsvFileListView) {

                 /*các ListCell là các phần tử con hay các hàng của list nó extends Labeled
                 nên có thể định dạng cho nó giống Labeled như màu sắc
                 ListCell không phải Interface nhưng ta vẫn tạo lớp ẩn danh kế thừa lớp này và
                 Override method updateItem của nó*/
                ListCell<ExcelFile> cell = new ListCell<ExcelFile>() {
                    @Override
                    protected void updateItem(ExcelFile csvFile, boolean empty) {
                        //vẫn giữ lại các cài đặt của lớp cha, chỉ cần sửa một vài giá trị
                        super.updateItem(csvFile, empty);

                        if (csvFile != null && !empty) {
                            // vùng chứa tên file
                            Label labelName = new Label(csvFile.getName());
                            // cài label có chiều ngang tối đa max
                            labelName.setMaxWidth(Double.MAX_VALUE);
                            // chọn chiều ngang tối thiểu là 188, do cài HBox.setHgrow(labelName, Priority.SOMETIMES);
                            // nên nếu hbox có thể dài hơn thì label cũng có thể dài theo
                            labelName.setPrefWidth(188);
                            labelName.setWrapText(true);
                            // chữ ở trung tâm
                            labelName.setAlignment(Pos.CENTER);
//                            labelName.setStyle("-fx-background-color: blue;");
                            labelName.setStyle("-fx-padding: 0 0 0 3;");
                            // chữ màu BLUE
                            labelName.setTextFill(Color.BLUE);
                            // cho label chiếm chiều ngang hết những vùng còn thừa
                            HBox.setHgrow(labelName, Priority.SOMETIMES);

                            // lấy tổng chiều dài bozai, tính theo m lên / 1000
                            double kouzaiChouGoukei = csvFile.getKouzaiChouGoukei() / 1000;
                            // lấy tổng chiều dài sản phẩm, tính theo m nên / 1000 và do chiều dài bị x100 trước nên cần / 100
                            double seiHinChouGoukei = (csvFile.getSeiHinChouGoukei() / 100) / 1000;

                            // tạo đối tượng fomat số theo định dạng Nhật
                            NumberFormat numberFormat = NumberFormat.getInstance(new Locale("ja", "JA"));

                            // làm tròn các chiều dài về 2 chữ số thập phân
                            // hàm round trả về long nên cần x100 trước, nó sẽ làm tròn số về long, sau đó chuyển sang double và
                            // /100 sẽ được số double có 2 số sau phần thập phân
                            String formattedKouzaiChouGoukei = numberFormat.format((double) Math.round(kouzaiChouGoukei * 100) / 100);
                            String formattedseiHinChouGoukei = numberFormat.format((double) Math.round(seiHinChouGoukei * 100) / 100);

                            // vùng chứa tổng chiều dài bozai
                            Label labelKouzaiChou = new Label(formattedKouzaiChouGoukei + "  ");
                            // nếu vùng chứa tổng chiều dài là 0(chưa phải file cuối cùng của vật liệu đang xét)
                            // thì cho nó hiển thị rỗng
                            if (kouzaiChouGoukei == 0) {
                                labelKouzaiChou.setText("_____  ");
                            }
                            labelKouzaiChou.setMinWidth(USE_PREF_SIZE);
                            labelKouzaiChou.setPrefWidth(59);
                            labelKouzaiChou.setStyle("-fx-text-fill: #F57C00");
                            labelKouzaiChou.setAlignment(Pos.CENTER_RIGHT);

                            // vùng chứa tổng chiều dài sản phẩm
                            Label labelSeihinChou = new Label("  " + formattedseiHinChouGoukei);

                            if (seiHinChouGoukei == 0) {
                                labelSeihinChou.setText("  _____");
                            }
                            labelSeihinChou.setMinWidth(USE_PREF_SIZE);
                            labelSeihinChou.setPrefWidth(59);
                            labelSeihinChou.setStyle("-fx-text-fill: #00796B");

                            // hbox sẽ chứa tất cả các control
                            HBox hBox = new HBox();
                            hBox.setAlignment(Pos.CENTER_RIGHT);
                            hBox.setMaxWidth(Double.MAX_VALUE);
                            hBox.setStyle("-fx-font-weight: bold; -fx-background-color: #DCEDC8; -fx-padding: 3 3 3 3");

                            // tạo luồng đọc file ảnh
                            Class<ConVertExcelAnd3BCController> clazz = ConVertExcelAnd3BCController.class;
                            InputStream input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/ok.png");

                            // lấy tên vật liệu của file
                            String koSyuName = csvFile.getKouSyuName();
                            // chọn ảnh dựa theo tên vật liệu
                            if (koSyuName.equalsIgnoreCase("[")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/U.png");
                            } else if (koSyuName.equalsIgnoreCase("C")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/C.png");
                            } else if (koSyuName.equalsIgnoreCase("K")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/P.png");
                            } else if (koSyuName.equalsIgnoreCase("L")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/L.png");
                            } else if (koSyuName.equalsIgnoreCase("H")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/H.png");
                            } else if (koSyuName.equalsIgnoreCase("FB")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/FB.png");
                            } else if (koSyuName.equalsIgnoreCase("CA")) {
                                input = clazz.getResourceAsStream("/com/lenha/excel_3bc_toriai/ICON/CA.png");
                            }

                            // control chứa ảnh
                            Image image;

                            try {
                                assert input != null;
                                // thêm luồng đọc ảnh vào control chứa ảnh
                                image = new Image(input);
                                // tạo ImageView chứa image để hiển thị ảnh
                                ImageView imageView = new ImageView(image);
                                imageView.setFitWidth(25);
                                imageView.setFitHeight(25);

                                // tạo vùng ngăn cách giữa 2 giá trị tổng chiều dài
                                Label separation = new Label("|");

                                // thêm các control vào hbox theo thứ tự xác định
                                hBox.getChildren().add(labelName);
                                hBox.getChildren().add(imageView);
                                hBox.getChildren().add(labelSeihinChou);
                                hBox.getChildren().add(separation);
                                hBox.getChildren().add(labelKouzaiChou);
//                                hBox.setSpacing(10);

                                // gán hbox cho dòng của list view
                                setGraphic(hBox);
                            } catch (NullPointerException e) {
                                System.out.println(e.getMessage());
                            }

                        } else {
                            setGraphic(null);
                        }
                    }
                };

                //trả về lớp ẩn danh kế thừa cell trên vừa Override lại các updateItem của nó
                return cell;
            }
        });

    }

    /**
     * đóng chương trình
     */
    public void closeApp(ActionEvent actionEvent) {
        Platform.exit();
        System.exit(0);
    }

    /**
     * mở dialog giới thiệu về chương trình
     */
    public void openAbout(ActionEvent actionEvent) {
        // tạo dialog với gắn liền với cửa sổ gốc của ứng dụng
        Dialog<Object> dialog = new Dialog<>();
        dialog.initOwner(menuBar.getScene().getWindow());// lấy window đang chạy

        dialog.setTitle("About");

        // cho phép thay đổi kích thước
        dialog.setResizable(true);
        FXMLLoader loader = new FXMLLoader();
        loader.setLocation(ConVertExcelAnd3BCController.class.getResource("/com/lenha/excel_3bc_toriai/about.fxml"));// thêm ui fxml

        try {
            dialog.getDialogPane().setContent(loader.load());// liên kết ui fxml vào dialog
        } catch (IOException e) {
            System.out.println("Couldn't load the dialog");
            e.printStackTrace();
        }

        // lấy controller của ui FXMLLoader
        AboutController controller = loader.getController();
        // gọi hàm init của controller và truyền đối tượng của chính controller cửa sổ đang hiển thị này và dialog của chính controller cửa sổ mới cho nó
        controller.init(this, dialog);

        // hiển thị dialog
        dialog.show();

//        // nếu là nút ok thì thêm item nhập từ dialog vào listview
//        if (result.isPresent() && result.get() == ButtonType.OK) {
//            DialogController controller = loader.getController();// lấy controller của ui fxml
//            TodoItem newItem = controller.processResult();// nhận TodoItem từ hàm của controller trả về, hàm này đã thêm item vòa list liên kết với listview
//            todoListView.getSelectionModel().select(newItem);//cho list view chọn item trên
//        } else {
//            System.out.println("cancel");
//        }
//        FXMLLoader loader = new FXMLLoader(this.getClass().getResource("About.fxml"));
//        Parent root = (Parent)loader.load();
//        Stage stage = new Stage();
//        stage.initOwner(menuBar.getScene().getWindow());
//        stage.setScene(new Scene(root));
//        stage.setTitle("Update a Contact");
//        stage.show();


    }

    /**
     * chuyển file excel máy cắt sang file 3bc
     *
     * @param actionEvent
     */
    @FXML
    public void convertExcelTo3bcFile(ActionEvent actionEvent) {
        // link file excel
        String excelFilePath;
        // link thư mục chứa file 3bc sẽ tạo
        String _3bcFileDirPath;

        // yêu cầu chọn địa chỉ file và thư mục khi 2 địa chỉ này chưa được chọn
        // nếu chọn xong thì phải chuyển dữ liệu thành công thì mới thoát được vòng lặp
        while (true) {
            // lấy link từ các ô đang hiển thị
            excelFilePath = linkExcelFile.getText();
            _3bcFileDirPath = link3bcDir.getText();

            // tạo file theo các link trên
            File excelFile = new File(excelFilePath);
            File _3bcFileDir = new File(_3bcFileDirPath);

            // kiểm tra hợp lệ địa chỉ file và thư mục
            boolean isFileExcel = excelFile.isFile();
            boolean isDir = _3bcFileDir.isDirectory();

            // nếu không phải là file excel thì yêu cầu chọn lại
            if (!isFileExcel) {
                // hiển thị alert yêu cầu chọn lại file excel
                confirmAlert.setTitle(CONFIRM_EXCEL_FILE_LINK_TITLE);
                confirmAlert.setHeaderText(CONFIRM_EXCEL_FILE_LINK_HEADER);
                confirmAlert.setContentText(CONFIRM_EXCEL_FILE_LINK_CONTENT);
                updateLangAlert(confirmAlert);
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu đồng ý thì gọi lại hàm chọn file
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File fileSelected = getExcelFile();
                    // nếu file đã chọn null nghĩa là người dùng click vào nút cancel khi chọn file
                    // thì thoát và không làm gì nữa
                    if (fileSelected == null) {
                        return;
                    }

                    /* sau khi đã chọn xong file EXCEL thì gán luôn file do hàm chọn trả về để hàm lấy file cho việc chuyển bên dưới
                     hoạt động đúng mà không phải thêm 1 vòng lặp nữa tính lại file */
                    excelFile = fileSelected;

                    // nếu file chọn không đúng thì nhảy sang sang vòng lặp mới và chọn lại từ đầu
                    if (!excelFile.isFile()) {
                        continue;
                    }
                }
                // nếu không đồng ý thì thoát hàm
                else {
                    return;
                }
            }

            // nếu không phải là thư mục thì yêu cầu chọn lại
            if (!isDir) {
                /* nếu địa chỉ file 3BC đã xác nhận thì nó sẽ tự động lấy địa chỉ thư mục chứa file 3BC đó nhập vào
                 link3bcDir, mà trước đó đã xác nhận chưa chọn thư mục chứa file đã chuyển nên cần xóa text của
                 link3bcDir đi để người dùng xác nhận lại, tránh hiển thị địa chỉ mặc định trên ỏ link3bcDir làm khó hiểu */
                link3bcDir.setText("");
                SetupData.getInstance().setLinkSave3bcFileDir("");
                // hiển thị alert yêu cầu chọn lại
                confirmAlert.setTitle(CONFIRM_3BC_FILE_DIR_TITLE);
                confirmAlert.setHeaderText(CONFIRM_3BC_FILE_DIR_HEADER);
                confirmAlert.setContentText(CONFIRM_3BC_FILE_DIR_CONTENT);
                updateLangAlert(confirmAlert);
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu là nút ok thì gọi hàm chọn thư mục
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File dirSelected = setSave3bcFileDir();

                    // nếu thư mục trả về null tức người dùng hủy chọn bằng nút cancel thì thoát hàm và không làm gì nữa
                    if (dirSelected == null) {
                        return;
                    }

                    /* sau khi đã chọn xong thư mục thì gán luôn thư mục do hàm chọn trả về để hàm lấy thư mục cho việc chuyển bên dưới
                     hoạt động đúng mà không phải thêm 1 vòng lặp nữa tính lại thư mục */
                    _3bcFileDir = dirSelected;

                    // nếu thư mục trả về không đúng thì nhảy sang sang vòng lặp mới và chọn lại từ đầu
                    if (!_3bcFileDir.isDirectory()) {
                        continue;
                    }
                }
                // nếu không đồng ý thì thoát hàm
                else {
                    return;
                }
            }

            // đến đây nếu không bị return thì đã chọn xong 2 địa chỉ file
            if (excelFile.isFile() && _3bcFileDir.isDirectory()) {
                System.out.println("đã chọn xong 2 địa chỉ");
                System.out.println("LINK EXCEL FILE: " + excelFile.getAbsolutePath());
                System.out.println("LINK 3BC DIR: " + _3bcFileDir.getAbsolutePath());
            }

            // gọi hàm chuyển file từ class static ReadPDFToExcel
            try {
                // biến kiểm tra trong danh sách các sheet tính vật liệu có sheet nào đó có vật liệu không giống với các vật liệu đã cài đặt sẵn trong chương trình không,
                // nếu có thì sẽ thay vật liệu của toàn bộ các sheet bằng bộ vật liệu tự cho trong danh sách dự phòng đã tạo khi khởi tạo chương trình
                // bộ vật liệu dự phòng lấy từ file excel VAT_LIEU_DU_PHONG.xlsx
                boolean co1VatLieuKhongTonTai = false;

                co1VatLieuKhongTonTai = ExcelTo3BC.convertExcelTo3bc(excelFile.getAbsolutePath(), _3bcFileDir.getAbsolutePath());
                if (co1VatLieuKhongTonTai){
                    // hiển thị alert cảnh báo đã đổi vật liệu không có trong danh sách sang bộ vật liệu dự phòng
                    confirmAlert.setAlertType(Alert.AlertType.WARNING);
                    confirmAlert.setTitle(CONFIRM_CONVERT_COMPLETE_TITLE);
                    confirmAlert.setHeaderText(CO_VAT_LIEU_KHONG_TON_TAI);
                    confirmAlert.setContentText(THONG_BAO_DOI_VAT_LIEU);
                    confirmAlert.showAndWait();

                }

                // hiển thị alert chuyển file thành công
                confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                confirmAlert.setTitle(CONFIRM_CONVERT_COMPLETE_TITLE);
                confirmAlert.setHeaderText(CONFIRM_CONVERT_EXCEL_TO_3BC_COMPLETE_HEADER);
                confirmAlert.setContentText(CONFIRM_CONVERT_EXCEL_TO_3BC_COMPLETE_CONTENT);

                updateLangAlert(confirmAlert);

                Optional<ButtonType> result = confirmAlert.showAndWait();

                // nếu là nút ok thì copy đường dẫn thư mục chứa file chl vào clipboard và mở thư mục này
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    copyContentToClipBoard(_3bcFileDir.getAbsolutePath());
                    // mở thư mục chứa file chl
                    Desktop.getDesktop().open(_3bcFileDir);
                }

                return;
            } catch (Exception e) {
                System.out.println(e.getMessage());
                e.printStackTrace();
                // xóa hết các button, đổi alert sang dạng error rồi thêm lại 2 nút ok và cancel
                confirmAlert.getButtonTypes().clear();
                confirmAlert.setAlertType(Alert.AlertType.ERROR);
                confirmAlert.getButtonTypes().add(ButtonType.CANCEL);
                confirmAlert.getButtonTypes().add(ButtonType.OK);

                confirmAlert.setTitle(ERROR_CONVERT_TITLE);
                confirmAlert.setHeaderText(ERROR_CONVERT_EXCEL_TO_3BC_HEADER);
                confirmAlert.setContentText(ERROR_CONVERT_CONTENT);
                // cập nhật ngôn ngữ cho alert
                updateLangAlert(confirmAlert);

                // nếu là sự kiện không ghi được file excel do file trùng tên với file sắp tạo đang được mở
                // thì in ra cảnh báo và thoát
                if (e instanceof FileNotFoundException) {
                    confirmAlert.getButtonTypes().clear();
                    confirmAlert.getButtonTypes().add(ButtonType.OK);
                    //NEW
                    confirmAlert.setHeaderText("Tên file 3BC đang tạo: (\"" + ReadPDFToExcel.fileName + "\") trùng tên với 1 file 3BC khác đang được mở nên không thể ghi đè");
                    confirmAlert.setContentText("Hãy đóng file 3BC đang mở để tiếp tục!");
                    System.out.println("File đang được mở bởi người dùng khác");
                    updateLangAlert(confirmAlert);
                    confirmAlert.showAndWait();

                    // chuyển lại alert về dạng confirm và thêm nút cancel
                    confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                    confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

                    return;
                }

                // 3BC KHÔNG CÓ LỖI NÀY NHƯ CHL
                /*// nếu là lỗi quá 99 dòng thì thông báo và thoát
                if (e instanceof TimeoutException) {
                    confirmAlert.getButtonTypes().clear();
                    confirmAlert.getButtonTypes().add(ButtonType.OK);
                    confirmAlert.setHeaderText("File CHL đang tạo: (\"" + ReadPDFToExcel.fileName + "\") trong một boZai duy nhất có số dòng sản phẩm cần ghi lớn hơn 99 nên không thể ghi");
                    confirmAlert.setContentText("Hãy chỉnh sửa lại dữ liệu vật liệu đang chuyển để tiếp tục!");
                    System.out.println("Vật liệu có số dòng lớn hơn 99");
                    updateLangAlert(confirmAlert);
                    confirmAlert.showAndWait();

                    // chuyển lại alert về dạng confirm và thêm nút cancel
                    confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);
                    confirmAlert.getButtonTypes().add(ButtonType.CANCEL);

                    return;
                }*/


                // nếu là lỗi ghi file thì thông báo
                Optional<ButtonType> result = confirmAlert.showAndWait();

                // chuyển lại alert về dạng confirm
                confirmAlert.setAlertType(Alert.AlertType.CONFIRMATION);

                // nếu chọn ok thì gọi lại hàm chọn file EXCEL để chọn file khác
                // nếu chọn cancel thì thoát
                if (result.isPresent() && result.get() == ButtonType.OK) {
                    File fileSelected2 = getExcelFile();

                    // nếu không chọn file thì thoát
                    if (fileSelected2 == null) {
                        return;
                    }
                } else {
                    return;
                }
            }
        }
    }

}