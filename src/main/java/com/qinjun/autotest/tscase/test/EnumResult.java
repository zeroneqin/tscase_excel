package com.qinjun.autotest.tscase.test;

import java.util.List;

public enum EnumResult {
    PASS("Pass"), FAIL("Fail"),SKIP("Skip");

    ThreadLocal<String>  resultMsg;
    ThreadLocal<List<String>> additionaInfos = new ThreadLocal<List<String>>();

    private EnumResult(final String resultMsg) {
        this.resultMsg=new ThreadLocal<String>() {
            protected String initaValue() {return resultMsg;}
        };
    }

    public ThreadLocal<String> getResultMsg() {
        return resultMsg;
    }

    public void setResultMsg(ThreadLocal<String> resultMsg) {
        this.resultMsg = resultMsg;
    }

    public List<String> getAdditionaInfos() {
        return additionaInfos.get();
    }

    public void setAdditionaInfos(List<String> additionaInfos) {
        this.additionaInfos.set(additionaInfos);
    }
}
