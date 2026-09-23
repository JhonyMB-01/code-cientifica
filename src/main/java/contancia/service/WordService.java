package contancia.service;

import contancia.config.DocumentConfig;
import contancia.processor.WordTemplateProcessor;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFFooter;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

@ApplicationScoped
public class WordService {

    @Inject
    DocumentConfig documentConfig;

    @Inject
    WordTemplateProcessor templateProcessor;


    public Path generarWord(
            Map<String, String> datos) {

        String templatePath =
                documentConfig.getWordTemplate();

        try (
                FileInputStream fis =
                        new FileInputStream(templatePath);

                XWPFDocument document =
                        new XWPFDocument(fis)
        ) {

            templateProcessor.procesarParrafos(
                    document.getParagraphs(),
                    datos
            );

            for (XWPFHeader header :
                    document.getHeaderList()) {

                templateProcessor.procesarParrafos(
                        header.getParagraphs(),
                        datos
                );
            }

            for (XWPFFooter footer :
                    document.getFooterList()) {

                templateProcessor.procesarParrafos(
                        footer.getParagraphs(),
                        datos
                );
            }

            Path directorio =
                    Files.createTempDirectory(
                            "document-word-"
                    );

            Path docx =
                    directorio.resolve(
                            "documento.docx"
                    );

            try (FileOutputStream fos =
                         new FileOutputStream(
                                 docx.toFile()
                         )) {

                document.write(fos);
            }

            return docx;

        } catch (Exception e) {

            throw new RuntimeException(
                    "No se pudo generar el Word",
                    e
            );
        }
    }
}
