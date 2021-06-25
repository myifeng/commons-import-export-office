package com.github.myifeng.commons.office.processor;


import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;

public interface SheetProcessor {

    void write(Map<String, String> map);

    void write(OutputStream out) throws IOException;

}
