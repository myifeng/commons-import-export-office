package com.github.myifeng.commons.office.processor.impl;

import com.github.myifeng.commons.office.processor.WordProcessor;
import org.apache.poi.hwpf.HWPFDocument;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Collection;
import java.util.stream.Stream;

/**
 * Processor the Microsoft Word 97(-2007) file by HWPFDocument.
 *
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
    public void writeEntity(Object object) {
        Stream.of(object.getClass().getDeclaredFields()).forEach(field -> {
            try {
                field.setAccessible(true);
                String name = field.getName();
                Object value = field.get(object);
                Class<?> type = field.getType();
                if (Collection.class.isAssignableFrom(type)) {
                    //TODO: Process table

                } else {
                    //TODO: Process text

                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        });
    }

    @Override
    public void write(OutputStream out) throws IOException {
        doc.write(out);
    }

}
