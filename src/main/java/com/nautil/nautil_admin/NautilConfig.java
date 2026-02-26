package com.nautil;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;

@Configuration
public class NautilConfig {

    @Value("${nautil.excel.verification-file}")
    private String verificationFilePath;

    @Value("${nautil.excel.template-file}")
    private String templateFilePath;

    @Value("${nautil.excel.output-dir}")
    private String outputDir;

    @Value("${nautil.excel.sheet-global}")
    private String sheetGlobal;

    @Value("${nautil.excel.sheet-tx1}")
    private String sheetTx1;

    @Value("${nautil.excel.sheet-tx2}")
    private String sheetTx2;

    @Value("${nautil.excel.sheet-tx3}")
    private String sheetTx3;

    @Value("${nautil.email.default-to}")
    private String defaultEmailTo;

    @Value("${nautil.email.from}")
    private String emailFrom;

    public String getVerificationFilePath() { return verificationFilePath; }
    public String getTemplateFilePath() { return templateFilePath; }
    public String getOutputDir() { return outputDir; }
    public String getSheetGlobal() { return sheetGlobal; }
    public String getSheetTx1() { return sheetTx1; }
    public String getSheetTx2() { return sheetTx2; }
    public String getSheetTx3() { return sheetTx3; }
    public String getDefaultEmailTo() { return defaultEmailTo; }
    public String getEmailFrom() { return emailFrom; }

    public void setVerificationFilePath(String verificationFilePath) { this.verificationFilePath = verificationFilePath; }
    public void setTemplateFilePath(String templateFilePath) { this.templateFilePath = templateFilePath; }
    public void setOutputDir(String outputDir) { this.outputDir = outputDir; }
}
