package com.qinjun.autotest.tscase;

import com.qinjun.autotest.tscase.util.ExcelUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Parameters;

public class Suite {
    private final static Logger logger = LoggerFactory.getLogger(Suite.class);
    protected static int caseNum = -1;
    protected String caseExcel;
    protected String caseSheet;
    protected String type;

    @Parameters({"caseExcel","caseSheet","type"})
    @BeforeSuite
    public void beforeSuite(String caseExcel,String caseSheet,String type) {
        this.type = type;
        this.caseExcel=caseExcel;
        this.caseSheet=caseSheet;
        try {
            Sheet sheet = ExcelUtil.getSheet(caseExcel,caseSheet);
            int lastRowNum = ExcelUtil.getLastRowNum(sheet);
            for (int i=1;i<lastRowNum;i++) {
                ExcelUtil.setCellValue(caseExcel,caseSheet,i,4,"");
                ExcelUtil.setCellValue(caseExcel,caseSheet,i,5,"");
            }
        }
        catch (Exception e) {
            logger.error("Get exception;"+e);
        }
        logger.info("==========[开始执行测试集]==========");
    }

    @AfterSuite
    public void afterSuite() {
        logger.info("==========[结束执行测试集]==========");
    }
}
