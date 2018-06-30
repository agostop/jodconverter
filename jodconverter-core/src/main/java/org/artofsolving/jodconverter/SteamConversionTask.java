//
// JODConverter - Java OpenDocument Converter
// Copyright 2004-2012 Mirko Nasato and contributors
//
// JODConverter is Open Source software, you can redistribute it and/or
// modify it under either (at your option) of the following licenses
//
// 1. The GNU Lesser General Public License v3 (or later)
//    -> http://www.gnu.org/licenses/lgpl-3.0.txt
// 2. The Apache License, Version 2.0
//    -> http://www.apache.org/licenses/LICENSE-2.0.txt
//
package org.artofsolving.jodconverter;

import static org.artofsolving.jodconverter.office.OfficeUtils.cast;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.OfficeException;

import com.sun.star.lang.XComponent;
import com.sun.star.lib.uno.adapter.ByteArrayToXInputStreamAdapter;
import com.sun.star.lib.uno.adapter.OutputStreamToXOutputStreamAdapter;
import com.sun.star.util.XRefreshable;

public class SteamConversionTask extends AbstractStreamConversionTask{

    private final DocumentFormat outputFormat;

    private Map<String,?> defaultLoadProperties;
    private DocumentFormat inputFormat;

    public SteamConversionTask(InputStream inputFile, OutputStream outputFile, DocumentFormat outputFormat) {
        super(inputFile, outputFile);
        this.outputFormat = outputFormat;
    }

    public void setDefaultLoadProperties(Map<String, ?> defaultLoadProperties) {
        this.defaultLoadProperties = defaultLoadProperties;
    }

    public void setInputFormat(DocumentFormat inputFormat) {
        this.inputFormat = inputFormat;
    }

    @Override
    protected void modifyDocument(XComponent document) throws OfficeException {
        XRefreshable refreshable = cast(XRefreshable.class, document);
        if (refreshable != null) {
            refreshable.refresh();
        }
    }

    @Override
    protected Map<String,?> getLoadProperties(InputStream inputStream) {
        Map<String,Object> loadProperties = new HashMap<String,Object>();
        if (defaultLoadProperties != null) {
            loadProperties.putAll(defaultLoadProperties);
        }
        if (inputFormat != null && inputFormat.getLoadProperties() != null) {
            loadProperties.putAll(inputFormat.getLoadProperties());
        }
        
        ByteArrayToXInputStreamAdapter byteArray = null;
        try {
        		byteArray = new ByteArrayToXInputStreamAdapter(IOUtils.toByteArray(inputStream));
		} catch (IOException e) {
			byteArray = null;
		}
		loadProperties.put("InputStream", byteArray);
        return loadProperties;
    }

    @Override
    protected Map<String,?> getStoreProperties(OutputStream outputStream, XComponent document) {
        DocumentFamily family = OfficeDocumentUtils.getDocumentFamily(document);
        
        Map<String,Object> storeProperties = new HashMap<String,Object>();
        storeProperties.putAll(outputFormat.getStoreProperties(family));
        storeProperties.put("OutputStream", new OutputStreamToXOutputStreamAdapter(outputStream));
        
        return storeProperties;
    }

}
