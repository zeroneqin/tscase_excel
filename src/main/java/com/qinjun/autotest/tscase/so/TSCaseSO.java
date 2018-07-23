package com.qinjun.autotest.tscase.so;

import com.qinjun.autotest.tscase.util.HttpResponse;
import com.qinjun.autotest.tscase.util.HttpUtil;
import org.apache.http.entity.ContentType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TSCaseSO {
    private final static Logger logger = LoggerFactory.getLogger(TSCaseSO.class);

    public static HttpResponse sendRequest(String url, String req) {
        HttpResponse httpResponse = null;
        String fullUrl = url;

        try {
            logger.info("url:"+fullUrl);
            httpResponse = HttpUtil.sendPost(fullUrl,null,req, ContentType.TEXT_PLAIN);
            if (httpResponse == null) {
                logger.error("Http response is null");
            }
        }
        catch (Exception e) {
            logger.error("Get exception:"+e);
        }

        return httpResponse;
    }
}
