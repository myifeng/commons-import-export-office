package com.github.myifeng.commons.office.factory;

import com.github.myifeng.commons.office.processor.WordProcessor;
import com.github.myifeng.commons.office.processor.ExcelProcessor;
import com.github.myifeng.commons.office.processor.impl.DocxProcessorImpl;
import com.github.myifeng.commons.office.processor.impl.DocProcessorImpl;
import com.google.common.io.ByteStreams;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;


/**
 * Processor Factory
 * <br/>
 * create a {@link ExcelProcessor} or {@link WordProcessor} instance.
 * @author myifeng
 */
public class ProcessorFactory {

    /**
     * create a {@link ExcelProcessor} instance.
     * @param inputStream
     * @return ExcelProcessor
     */
    public static ExcelProcessor createExcelProcessor(InputStream inputStream) {

        return null;
    }

    /**
     * create a {@link WordProcessor} instance.
     * @param inputStream
     * @return WordProcessor
     */
    public static WordProcessor createWordProcessor(InputStream inputStream) {
        try {
            byte[] bytes = ByteStreams.toByteArray(inputStream);
            ByteArrayInputStream stream = new ByteArrayInputStream(bytes);
            try {
                return new DocProcessorImpl(new HWPFDocument(stream));
            } catch (OfficeXmlFileException e) {
                return new DocxProcessorImpl(new XWPFDocument(stream));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
