package com.github.myifeng.commons.office.processor;


import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;

/**
 * Processor Excel file.
 * @author myifeng
 */
public interface ExcelProcessor {

    void write(Map<String, String> map);

    void write(OutputStream out) throws IOException;

}
