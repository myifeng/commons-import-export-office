package com.github.myifeng.commons.office.processor.impl;

import com.github.myifeng.commons.office.processor.DocumentProcessor;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.IOException;
import java.io.OutputStream;

public class DocumentProcessorImpl implements DocumentProcessor {
	
	private HWPFDocument doc;

	public DocumentProcessorImpl(HWPFDocument doc) {
		this.doc = doc;
	}
	
	@Override
	public void replaceText(String src, String dest) {
		doc.getRange().replaceText(src, dest);
	}

	@Override
	public void write(OutputStream out) throws IOException {
		doc.write(out);
	}
}
