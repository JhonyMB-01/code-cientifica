package contancia.service;

import contancia.config.DocumentConfig;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.TimeUnit;

@ApplicationScoped
public class PdfService {

    @Inject
    DocumentConfig documentConfig;

    public byte[] convertirDocxAPdf(
            Path docxPath) throws Exception {

        Path outputDir =
                Files.createTempDirectory("document-pdf-");

        try {

            Path soffice =
                    Path.of(documentConfig.getLibreOfficePath());

            ProcessBuilder processBuilder =
                    new ProcessBuilder(
                            soffice.toString(),

                            "--headless",
                            "--convert-to",
                            "pdf:writer_pdf_Export",

                            "--outdir",
                            outputDir.toString(),

                            docxPath.toString()
                    );

            processBuilder
                    .redirectErrorStream(true);

            Process process =
                    processBuilder.start();

            String consola =
                    new String(
                            process.getInputStream()
                                    .readAllBytes()
                    );

            boolean terminado =
                    process.waitFor(
                            60,
                            TimeUnit.SECONDS
                    );

            if (!terminado) {

                process.destroyForcibly();

                throw new RuntimeException(
                        "LibreOffice excedió el tiempo máximo de conversión"
                );
            }

            if (process.exitValue() != 0) {

                throw new RuntimeException(
                        "Error convirtiendo DOCX a PDF. " +
                                "Código: " +
                                process.exitValue() +
                                ". Salida: " +
                                consola
                );
            }

            /*
             * LibreOffice genera el PDF con el mismo
             * nombre base del DOCX.
             */
            String nombrePdf =
                    docxPath.getFileName()
                            .toString()
                            .replaceFirst(
                                    "(?i)\\.docx$",
                                    ".pdf"
                            );

            Path pdfPath =
                    outputDir.resolve(nombrePdf);

            if (!Files.exists(pdfPath)) {

                throw new RuntimeException(
                        "LibreOffice terminó correctamente, " +
                                "pero no se encontró el PDF generado"
                );
            }

            return Files.readAllBytes(pdfPath);

        } finally {

            eliminarDirectorio(outputDir);
        }
    }


    private void eliminarDirectorio(
            Path directorio) {

        try {

            if (directorio == null ||
                    !Files.exists(directorio)) {

                return;
            }

            try (
                    var stream =
                            Files.walk(directorio)
            ) {

                stream
                        .sorted(
                                java.util.Comparator
                                        .reverseOrder()
                        )
                        .forEach(path -> {

                            try {
                                Files.deleteIfExists(path);
                            } catch (IOException ignored) {
                            }

                        });
            }

        } catch (Exception ignored) {
        }
    }
}
