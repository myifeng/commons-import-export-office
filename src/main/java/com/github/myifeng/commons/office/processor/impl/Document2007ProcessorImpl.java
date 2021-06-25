package com.github.myifeng.commons.office.processor.impl;

import com.github.myifeng.commons.office.processor.DocumentProcessor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.IOException;
import java.io.OutputStream;

public class Document2007ProcessorImpl implements DocumentProcessor {

	private XWPFDocument doc;

	public Document2007ProcessorImpl(XWPFDocument doc) {
		this.doc = doc;
	}

	@Override
	public void replaceText(String src, String dest) {
		for (XWPFParagraph p : doc.getParagraphs()) {
			for (XWPFRun r : p.getRuns()) {
				String text = r.toString();
				if (text.contains(src)) {
					text = text.replace(src, dest);
					r.setText(text);
				}
			}
		}
	}

	@Override
	public void write(OutputStream out) throws IOException {
		doc.write(out);
	}
}
