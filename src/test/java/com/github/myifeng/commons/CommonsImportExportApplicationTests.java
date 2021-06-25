package com.github.myifeng.commons;

import com.github.myifeng.commons.office.factory.ProcessorFactory;
import com.github.myifeng.commons.office.processor.DocumentProcessor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;

@SpringBootTest
class CommonsImportExportApplicationTests {

    private Path filePath = Paths.get(".\\test\\test.docx");

    @Test
    void testExportWordByReplaceText() throws IOException {
        //Blank Document
        XWPFDocument document = new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(filePath.toFile());
        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("{{SRC1}}");
        document.write(out);
        out.close();

        InputStream docInputStream = new FileInputStream(filePath.toFile());
        DocumentProcessor processor = ProcessorFactory.createDocumentProcessor(docInputStream);
        processor.replaceText("{{SRC1}}", "DEST1");
        processor.write(new FileOutputStream(".\\test\\out.docx"));
        XWPFDocument doc = new XWPFDocument(new FileInputStream(Paths.get(".\\test\\out.docx").toFile()));
        XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
        assert "DEST1".equals(extractor.getText());
    }

    @AfterEach
    void deleteFile() {
        filePath.toFile().getParentFile().delete();
    }

}
