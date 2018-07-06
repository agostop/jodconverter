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

import static org.artofsolving.jodconverter.office.OfficeUtils.SERVICE_DESKTOP;
import static org.artofsolving.jodconverter.office.OfficeUtils.cast;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUnoProperties;
import static org.artofsolving.jodconverter.office.OfficeUtils.toUrl;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

import org.artofsolving.jodconverter.office.OfficeContext;
import org.artofsolving.jodconverter.office.OfficeException;
import org.artofsolving.jodconverter.office.OfficeTask;
import org.artofsolving.jodconverter.office.OfficeUtils;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XStorable;
import com.sun.star.frame.XStorable2;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XComponent;
import com.sun.star.task.ErrorCodeIOException;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.CloseVetoException;
import com.sun.star.util.XCloseable;

public abstract class AbstractStreamConversionTask implements OfficeTask {

    private final InputStream inputStream;
    private final OutputStream outputStream;

    public AbstractStreamConversionTask(InputStream inputFile, OutputStream outputFile) {
        this.inputStream = inputFile;
        this.outputStream = outputFile;
    }

    protected abstract Map<String,?> getLoadProperties(InputStream inputStream);

    protected abstract Map<String,?> getStoreProperties(OutputStream outputStream, XComponent document);

    public void execute(OfficeContext context) throws OfficeException {
        XComponent document = null;
        try {
            document = loadDocument(context, inputStream);
            modifyDocument(document);
            storeDocument(document, outputStream);
        } catch (OfficeException officeException) {
            throw officeException;
        } catch (Exception exception) {
            throw new OfficeException("conversion failed", exception);
        } finally {
            if (document != null) {
                XCloseable closeable = cast(XCloseable.class, document);
                if (closeable != null) {
                    try {
                        closeable.close(true);
                    } catch (CloseVetoException closeVetoException) {
                        // whoever raised the veto should close the document
                    }
                } else {
                    document.dispose();
                }
            }
        }
    }

    private XComponent loadDocument(OfficeContext context, InputStream inputStream) throws OfficeException {
        if (inputStream == null) {
            throw new OfficeException("input document not found");
        }
        XComponentLoader loader = cast(XComponentLoader.class, context.getService(SERVICE_DESKTOP));
        Map<String,?> loadProperties = getLoadProperties(inputStream);
        XComponent document = null;
        try {
          //  document = loader.loadComponentFromURL(toUrl(inputStream), "_blank", 0, toUnoProperties(loadProperties));
        	document = loader.loadComponentFromURL("private:stream", "_blank", 0, toUnoProperties(loadProperties)); 
        } catch (IllegalArgumentException illegalArgumentException) {
            throw new OfficeException("could not load document", illegalArgumentException);
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not load document, errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not load document: ", ioException);
        }
        if (document == null) {
            throw new OfficeException("could not load document stream. ");
        }
        return document;
    }

    /**
     * Override to modify the document after it has been loaded and before it gets
     * saved in the new format.
     * <p>
     * Does nothing by default.
     * 
     * @param document
     * @throws OfficeException
     */
    protected void modifyDocument(XComponent document) throws OfficeException {
    	// noop
    }

    private void storeDocument(XComponent document, OutputStream outputStream) throws OfficeException {
        Map<String,?> storeProperties = getStoreProperties(outputStream, document);
        if (storeProperties == null) {
            throw new OfficeException("unsupported conversion");
        }
        
        try {
            cast(XStorable.class, document).storeToURL("private:stream", toUnoProperties(storeProperties));
        } catch (ErrorCodeIOException errorCodeIOException) {
            throw new OfficeException("could not store document, errorCode: " + errorCodeIOException.ErrCode, errorCodeIOException);
        } catch (IOException ioException) {
            throw new OfficeException("could not store document: " , ioException);
		}
    }

}
