package com.github.myifeng.commons.office.factory;

import com.github.myifeng.commons.office.processor.DocumentProcessor;
import com.github.myifeng.commons.office.processor.SheetProcessor;
import com.github.myifeng.commons.office.processor.impl.Document2007ProcessorImpl;
import com.github.myifeng.commons.office.processor.impl.DocumentProcessorImpl;
import com.google.common.io.ByteStreams;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;


public class ProcessorFactory {

    public static SheetProcessor createSheetProcessor(InputStream sheetInputStream) {

        return null;
    }

    public static DocumentProcessor createDocumentProcessor(InputStream documentInputStream) {
        try {
            byte[] bytes = ByteStreams.toByteArray(documentInputStream);
            ByteArrayInputStream stream = new ByteArrayInputStream(bytes);
            try {
                return new DocumentProcessorImpl(new HWPFDocument(stream));
            } catch (OfficeXmlFileException e) {
                return new Document2007ProcessorImpl(new XWPFDocument(stream));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
