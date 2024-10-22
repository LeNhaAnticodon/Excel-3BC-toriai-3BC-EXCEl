package com.lenha.excel_3bc_toriai;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.Node;
import javafx.scene.control.Button;
import javafx.scene.control.ButtonType;
import javafx.scene.control.Dialog;
import javafx.scene.control.Label;

public class AboutController {
    @FXML
    public Button okBtn;// nút đóng dialog
    @FXML
    public Label introduce;// text giới thiệu
    @FXML
    public Label introduceContent;// nội dung giới thiệu
    @FXML
    public Label using;// hướng dẫn sử dụng
    @FXML
    public Label usingContent;// nội dung hướng dẫn sử dụng
    @FXML
    public Label creator;// text tác giả

    private Dialog<Object> dialog;// dialog của cửa sổ

    /**
     * xử lý sự kiện khi click vào nút ok thì đóng dialog
     */
    @FXML
    public void okAbout(ActionEvent actionEvent) {
        dialog.setResult(Boolean.TRUE);
        dialog.close();
    }

    /**
     * khởi tạo dialog
     * tạo sự kiện cho nút x để đóng dialog
     * thêm các control của dialog này vào list để set ngôn ngữ cho các control khi ấn nút chuyển ngôn ngữ hoặc khi dialog bắt đầu hiển thị
     * @param conVertPdfToExcelCHLController controller của cửa sổ convert
     * @param dialog đối tượng dialog của chính cửa sổ này
     */
    public void init(ConVertPdfToExcelCHLController conVertPdfToExcelCHLController, Dialog<Object> dialog) {
        this.dialog = dialog;

        // đóng dialog bằng nút X
        // cần tạo nút close ẩn
        dialog.getDialogPane().getButtonTypes().add(ButtonType.CLOSE);
        Node closeButton = dialog.getDialogPane().lookupButton(ButtonType.CLOSE);
        closeButton.managedProperty().bind(closeButton.visibleProperty());
        closeButton.setVisible(false);

        // thêm các control của dialog này vào list để set ngôn ngữ cho các control
        ObservableList<Object> controls = FXCollections.observableArrayList(okBtn, introduce, introduceContent, using, usingContent, creator, dialog);
        // gọi hàm chuyển ngôn ngữ để hiển thị ngôn ngữ của các controls theo ngôn ngữ đã chọn
        conVertPdfToExcelCHLController.updateLangInBackground(conVertPdfToExcelCHLController.languages.getSelectedToggle(), controls);
    }
}
