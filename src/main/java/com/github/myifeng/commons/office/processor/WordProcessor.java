package com.github.myifeng.commons.office.processor;

import java.io.IOException;
import java.io.OutputStream;

/**
 * Processor Word file.
 * @author myifeng
 */
public interface WordProcessor {

    /**
     * Replace (all instances of) a piece of text with another...
     * @param src
     * The text to be replaced (e.g., "${organization}")
     * @param dest
     * The replacement text (e.g., "Apache Software Foundation")
     */
    void replaceText(String src, String dest);

    /**
     * Write out this document to an OutputStream
     * @param out
     * the java OutputStream you wish to write the file to
     * @throws IOException
     */
    void write(OutputStream out) throws IOException;

}
