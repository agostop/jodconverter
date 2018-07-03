package org.artofsolving.jodconverter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.io.IOUtils;
import org.artfosolving.jodconverter.util.ConverterUtils;
import org.artofsolving.jodconverter.office.OfficeManager;
import org.artofsolving.jodconverter.stream.OOoInputStream;
import org.artofsolving.jodconverter.stream.OOoOutputStream;

import com.sun.star.lib.uno.adapter.XOutputStreamToByteArrayAdapter;

public class OfficeToPdf {
	
	OfficeManager officeManager;
	
	public OfficeToPdf() {
		// TODO Auto-generated constructor stub
	}
	
	public static void main(String[] args) {
		ConverterUtils converterUtils = new ConverterUtils();
		converterUtils.initOfficeManager();
		OfficeDocumentConverter converter = converterUtils.getDocumentConverter();
		
		File file = new File("E:\\tmp\\CiwiZlsrVUOASwMtAHifSM9kI6s71.pptx");
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
			IOUtils.copy(new FileInputStream(file), bos);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		OOoInputStream ooInputStream = new OOoInputStream(bos.toByteArray());
		OOoOutputStream outputStream = new OOoOutputStream();
		
		converter.convert(ooInputStream, outputStream);
		
		File test = new File("E:\\tmp\\converted_tmpfile.pdf");
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(test);
			fos.write(outputStream.toByteArray());
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if (fos != null)
				try {
					fos.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
		
	}

}
