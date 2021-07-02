package com.github.myifeng.commons.office.processor.impl;

import com.github.myifeng.commons.office.processor.WordProcessor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.hwpf.usermodel.Range;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
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
        Stream.iterate(0, n -> n + 1)
                .limit(doc.getBookmarks().getBookmarksCount())
                .forEach(i -> {
                    try {
                        Bookmark bookmark = doc.getBookmarks().getBookmark(i);
                        Field field = object.getClass().getDeclaredField(bookmark.getName());
                        field.setAccessible(true);
                        Object value = field.get(object);
                        if (value != null) {
                            Range range = new Range(bookmark.getStart(), bookmark.getEnd(), doc);
                            if (range.getParagraph(0).isInTable()) {
                                //TODO: Process table

                            } else {
                                //TODO: Process text, convert data type
                                range.replaceText(value.toString(), false);
                            }
                        }
                    } catch (NoSuchFieldException e) {
                        //No field, pass!
                        e.printStackTrace();
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
