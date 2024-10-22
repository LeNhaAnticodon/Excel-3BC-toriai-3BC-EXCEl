package com.lenha.excel_3bc_toriai.model;

public class Setup {
    private String linkExcelFile = "";
    private String link3bcToriaiFile = "";
    private String linkSave3BCFileDir = "";
    private String lang = "";

    public String getLinkExcelFile() {
        return linkExcelFile;
    }

    public void setLinkExcelFile(String linkExcelFile) {
        this.linkExcelFile = linkExcelFile;
    }

    public String getLink3bcToriaiFile() {
        return link3bcToriaiFile;
    }

    public void setLink3bcToriaiFile(String link3bcToriaiFile) {
        this.link3bcToriaiFile = link3bcToriaiFile;
    }

    public String getLang() {
        return lang;
    }

    public void setLang(String lang) {
        this.lang = lang;
    }

    public String getLinkSave3BCFileDir() {
        return linkSave3BCFileDir;
    }

    public void setLinkSave3BCFileDir(String linkSave3BCFileDir) {
        this.linkSave3BCFileDir = linkSave3BCFileDir;
    }
}
