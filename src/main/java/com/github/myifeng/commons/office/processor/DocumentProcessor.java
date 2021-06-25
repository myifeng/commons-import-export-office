package com.github.myifeng.commons.office.processor;

import java.io.IOException;
import java.io.OutputStream;

public interface DocumentProcessor {

    void replaceText(String src, String dest);

    void write(OutputStream out) throws IOException;

}
