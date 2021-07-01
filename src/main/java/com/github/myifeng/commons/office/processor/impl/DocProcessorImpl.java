package com.github.myifeng.commons.office.processor.impl;

import com.github.myifeng.commons.office.processor.WordProcessor;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.IOException;
import java.io.OutputStream;

/**
 * Processor the Microsoft Word 97(-2007) file by HWPFDocument.
 * @author myifeng
 */
public class DocProcessorImpl implements WordProcessor {
	
	private HWPFDocument doc;

	public DocProcessorImpl(HWPFDocument doc) {
		this.doc = doc;
	}
	
	@Override
	public void replaceText(String src, String dest) {
		doc.getRange().replaceText(src, dest);
	}

	@Override
	public void replaceText(Object object) {
		//TODO: Get all the fields of the object
		//TODO: Get field type and value by reflect
		//TODO: Process text
	}

	@Override
	public void write(OutputStream out) throws IOException {
		doc.write(out);
	}

}
