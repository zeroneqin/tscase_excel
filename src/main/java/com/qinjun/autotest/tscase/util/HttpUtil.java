package com.qinjun.autotest.tscase.util;

import com.qinjun.autotest.tscase.exception.ExceptionUtil;
import com.qinjun.autotest.tscase.exception.TSCaseException;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.Header;
import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.entity.ContentType;
import org.apache.http.entity.StringEntity;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.util.EntityUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.util.Map;

public class HttpUtil {
    private final static Logger logger = LoggerFactory.getLogger(HttpUtil.class);

    public static HttpResponse sendPost(String url, Map<String,String> headerMap, String requestBody, ContentType contentType) throws TSCaseException {
        logger.info("Start to send post request");
        HttpResponse httpResponse = new HttpResponse();
        CloseableHttpClient httpClient = HttpClients.createDefault();

        try {
            HttpPost httpPost = new HttpPost();
            logger.info("Request:"+httpPost.getRequestLine());
            if (headerMap!=null) {
                for (String headerKey: headerMap.keySet()) {
                    String headerValue = headerMap.get(headerKey);
                    if (!StringUtils.isEmpty(headerValue)) {
                        logger.debug("Request header name:[{}], value:[{}]",headerKey,headerValue);
                        httpPost.addHeader(headerKey,headerValue);
                    }
                }
            }

            if (!StringUtils.isEmpty(requestBody)) {
                StringEntity stringRequestEntity = new StringEntity(requestBody,contentType);
                httpPost.setEntity(stringRequestEntity);
                logger.debug("Request content type:"+contentType);
                logger.debug("Request body:"+requestBody);
            }

            CloseableHttpResponse closeableHttpResponse = httpClient.execute(httpPost);
            Header[] headers = httpResponse.getHeaders();
            for (Header header: headers) {
                logger.debug("Response header name:[{}], value:[{}]",header.getName(),header.getValue());
            }

            httpResponse.setHeaders(headers);
            HttpEntity httpResponseEntity = closeableHttpResponse.getEntity();
            int status = closeableHttpResponse.getStatusLine().getStatusCode();
            httpResponse.setStatus(status);
            logger.info("Response status code:"+status);
            if (httpResponseEntity!=null) {
                logger.info("Response body length:"+httpResponseEntity.getContentLength());
                byte[] reponseByteBody = EntityUtils.toByteArray(httpResponseEntity);
                String responseBody = new String(reponseByteBody,"utf8");
                logger.info("Response body:"+responseBody);
                EntityUtils.consume(httpResponseEntity);
                httpResponse.setBody(responseBody);
                httpResponse.setByteBody(reponseByteBody);
            }
        }
        catch(Exception e) {
            String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(e);
            logger.error("Get exception:"+exceptionStackTrace);
            throw new TSCaseException(exceptionStackTrace);
        }
        finally {
            try {
                httpClient.close();
            }
            catch (IOException ioe) {
                String exceptionStackTrace = ExceptionUtil.getExceptionStackTrace(ioe);
                logger.error("Get exception:"+exceptionStackTrace);
                throw new TSCaseException(exceptionStackTrace);
            }
        }
        return httpResponse;
    }
}
