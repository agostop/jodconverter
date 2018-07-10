package org.artfosolving.jodconverter.util;

import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Map;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

import com.sun.star.document.UpdateDocMode;

public class ConverterUtils {

//    OpenOfficeConnection connection;
    OfficeManager officeManager;

    public void initOfficeManager() {
        ////            connection = new SocketOpenOfficeConnection(host,8100);
////            connection.connect();
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
        configuration.setOfficeHome("C:\\Program Files\\LibreOffice 5\\");
        configuration.setPortNumber(8100);
        officeManager = configuration.buildOfficeManager();
        officeManager.start();
        // 设置任务执行超时为5分钟
        // configuration.setTaskExecutionTimeout(1000 * 60 * 5L);//
        // 设置任务队列超时为24小时
        // configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);//
    }

    public OfficeDocumentConverter getDocumentConverter() {
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager, new ControlDocumentFormatRegistry());
        converter.setDefaultLoadProperties(getLoadProperties());
        return converter;
    }

    private Map<String,?> getLoadProperties() {
        Map<String,Object> loadProperties = new HashMap<String, Object>(10);
        loadProperties.put("Hidden", true);
        loadProperties.put("ReadOnly", true);
        loadProperties.put("UpdateDocMode", UpdateDocMode.QUIET_UPDATE);
        loadProperties.put("CharacterSet", Charset.forName("UTF-8").name());
        return loadProperties;
    }

    public void destroyOfficeManager(){
        if (null != officeManager && officeManager.isRunning()) {
            officeManager.stop();
        }
    }

}