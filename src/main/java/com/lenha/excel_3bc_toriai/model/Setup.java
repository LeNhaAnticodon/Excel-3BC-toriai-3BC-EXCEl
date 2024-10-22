package com.lenha.excel_3bc_toriai.model;

public class Setup {
    private String linkPdfFile = "";
    private String linkSaveCvsFileDir = "";
    private String lang = "";

    public String getLinkPdfFile() {
        return linkPdfFile;
    }

    public void setLinkPdfFile(String linkPdfFile) {
        this.linkPdfFile = linkPdfFile;
    }

    public String getLinkSaveCvsFileDir() {
        return linkSaveCvsFileDir;
    }

    public void setLinkSaveCvsFileDir(String linkSaveCvsFileDir) {
        this.linkSaveCvsFileDir = linkSaveCvsFileDir;
    }

    public String getLang() {
        return lang;
    }

    public void setLang(String lang) {
        this.lang = lang;
    }
}
