package contancia.service;

import Enmienda.model.DocumentoImagen;
import contancia.config.DocumentConfig;
import contancia.processor.WordImageProcessor;
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
import java.util.List;
import java.util.Map;

@ApplicationScoped
public class WordService {

    @Inject
    WordTemplateProcessor templateProcessor;

    @Inject
    WordImageProcessor imageProcessor;


    public Path generarWord(
            Map<String, String> datos, String path, List<DocumentoImagen> imagenes) {


        try (
                FileInputStream fis =
                        new FileInputStream(path);

                XWPFDocument document =
                        new XWPFDocument(fis)
        ) {
            /*
             * 1. Procesar texto
             */
            templateProcessor.procesarParrafos(
                    document.getParagraphs(),
                    datos
            );

            /*
             * 2. Procesar imágenes
             *
             * Solo Enmienda enviará imágenes.
             */
            if (imagenes != null &&
                    !imagenes.isEmpty()) {

                imageProcessor.procesarImagenes(
                        document,
                        imagenes
                );
            }

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
