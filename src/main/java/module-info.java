module com.example.convert_toriai_from_pdf_to_excel {
    requires javafx.controls;
    requires javafx.fxml;
    requires com.opencsv;
    requires org.apache.pdfbox;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.lenha.excel_3bc_toriai to javafx.fxml;
    exports com.lenha.excel_3bc_toriai;
}