package com.github.myifeng.commons;

import com.github.myifeng.commons.office.factory.ProcessorFactory;
import com.github.myifeng.commons.office.processor.WordProcessor;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.nio.file.Paths;

@SpringBootTest
class CommonsImportExportApplicationTests {

    private String testPath = ".test";

    @BeforeEach
    void initDir(){
        Paths.get(testPath).toFile().mkdirs();
    }

    @Test
    void testExportWordByReplaceText() throws IOException {
        WordProcessor processor = ProcessorFactory.createWordProcessor(this.getClass().getResourceAsStream("/exportByReplaceTextTemplate.doc"));
        processor.replaceText("{{SRC1}}", "DEST1");
        processor.replaceText("SRC2", "DEST2");

        String outPath = Paths.get(testPath,"out.doc").toString();
        processor.write(new FileOutputStream(outPath));
        HWPFDocument doc = new HWPFDocument(new FileInputStream(outPath));
        WordExtractor extractor = new WordExtractor(doc);
        String text = extractor.getText();

        assert !text.contains("{{SRC1}}");
        assert text.contains("DEST1");
        assert !text.contains("SRC2");
        assert text.contains("DEST2");
    }

    @AfterEach
    void deleteDir(){
        deleteFile(Paths.get(testPath).toFile());
    }

    private boolean deleteFile(File dirFile) {
        if (!dirFile.exists()) {
            return false;
        }
        if (dirFile.isFile()) {
            return dirFile.delete();
        } else {
            for (File file : dirFile.listFiles()) {
                deleteFile(file);
            }
        }
        return dirFile.delete();
    }

}
