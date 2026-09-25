package contancia.service;
import contancia.DocumentNotFoundException;
import contancia.config.DocumentConfig;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

@ApplicationScoped
public class DocumentService {

    @Inject
    ExcelService excelService;

    @Inject
    WordService wordService;

    @Inject
    PdfService pdfService;

    @Inject
    DocumentConfig documentConfig;

    /**
     * Genera el documento Word final.
     */
    public byte[] generarWord(
            String codigo) throws Exception {

        /*
         * 1. Obtener datos desde Excel.
         */
        Map<String, String> datos =
                obtenerDatos(codigo);


        /*
         * 2. Generar DOCX.
         */
        Path docx =
                null;

        try {

            docx =
                    wordService.generarWord(
                            datos, documentConfig.getWordTemplate(), null
                    );

            return Files.readAllBytes(
                    docx
            );

        } finally {

            eliminarTemporal(
                    docx
            );
        }
    }


    /**
     * Genera el documento PDF final.
     */
    public byte[] generarPdf(
            String codigo) throws Exception {

        /*
         * 1. Obtener datos desde Excel.
         */
        Map<String, String> datos =
                obtenerDatos(codigo);


        Path docx = null;

        try {

            /*
             * 2. Generar DOCX.
             */
            docx =
                    wordService.generarWord(
                            datos, documentConfig.getWordTemplate(), null
                    );


            /*
             * 3. Convertir DOCX → PDF.
             */
            return pdfService.convertirDocxAPdf(
                    docx
            );

        } finally {

            eliminarTemporal(
                    docx
            );
        }
    }


    private Map<String, String> obtenerDatos(
            String codigo) {

        Map<String, String> datos =
                excelService.buscarDatos(
                        codigo
                );

        if (datos == null) {

            throw new DocumentNotFoundException(
                    "Código no encontrado en Excel: " +
                            codigo
            );
        }

        return datos;
    }


    private void eliminarTemporal(
            Path archivo) {

        if (archivo == null) {
            return;
        }

        try {

            Path directorio =
                    archivo.getParent();

            Files.deleteIfExists(
                    archivo
            );

            if (directorio != null) {

                Files.deleteIfExists(
                        directorio
                );
            }

        } catch (Exception ignored) {
            // No interrumpir la respuesta por
            // un problema de limpieza temporal.
        }
    }
}
