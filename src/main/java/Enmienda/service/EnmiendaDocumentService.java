package Enmienda.service;

import Enmienda.model.EnmiendaExcelData;
import contancia.DocumentNotFoundException;
import contancia.config.DocumentConfig;
import contancia.service.PdfService;
import contancia.service.WordService;
import extencion.dto.GenerarDocumentoRequest;
import extencion.dto.GenerarPdfResponse;
import extencion.model.ExtensionExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import utils.GenerarWordResponse;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.mendoza.constants.Constantes.*;

@ApplicationScoped
public class EnmiendaDocumentService {

    @Inject
    EnmiendaExcelService enmiendaExcelService;

    @Inject
    DocumentConfig documentConfig;

    @Inject
    WordService wordService;

    @Inject
    PdfService pdfService;

    public GenerarWordResponse generarWordEnmienda(GenerarDocumentoRequest request)
            throws IOException {

        validarRequest(request);

        if (!"ENMIENDA".equalsIgnoreCase(
                request.getTipoDocumento())) {

            throw new IllegalArgumentException(
                    "Por el momento solamente se encuentra "
                            + "implementado el documento ENMIENDA"
            );
        }

        Path excelPath = Path.of(documentConfig.getBasePath())
                .resolve(documentConfig.getExcelFile());

        EnmiendaExcelData data =
                enmiendaExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Enmienda"
            );
        }

        Path plantilla =
                Path.of(documentConfig.getBasePath())
                        .resolve(documentConfig.getPlantillasPath())
                        .resolve(documentConfig.getPlantillaEnmienda());

        /*
         * 2. Generar DOCX.
         */
        Path docx =
                null;

        try {

            docx = wordService.generarWord(
                    construirValores(data), plantilla.toString(), null
            );

            return new GenerarWordResponse(Files.readAllBytes(docx),
                    String.format(ENMIENDA_WORD_NOMBRE, data.getCodigo(), data.getConstancia()));

        } finally {

            eliminarTemporal(
                    docx
            );
        }
    }

    /**
     * Genera el documento PDF final.
     */
    public GenerarPdfResponse generarPdf(
            GenerarDocumentoRequest request) throws Exception {

        /*
         * 1. Obtener datos desde Excel.
         */
        Path excelPath = Path.of(documentConfig.getBasePath())
                .resolve(documentConfig.getExcelFile());

        EnmiendaExcelData data =
                enmiendaExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Enmienda"
            );
        }

        Path plantilla =
                Path.of(documentConfig.getBasePath())
                        .resolve(documentConfig.getPlantillasPath())
                        .resolve(documentConfig.getPlantillaEnmienda());


        Path docx = null;

        try {

            /*
             * 2. Generar DOCX.
             */
            docx =
                    wordService.generarWord(
                            construirValores(data), plantilla.toString(), null
                    );


            /*
             * 3. Convertir DOCX → PDF.
             */
            return new GenerarPdfResponse(
                    pdfService.convertirDocxAPdf(docx),
                    String.format(
                            ENMIENDA_PDF_NOMBRE,
                            data.getCodigo(),
                            data.getConstancia()
                    ));


        } finally {

            eliminarTemporal(
                    docx
            );
        }
    }

    private void validarRequest(
            GenerarDocumentoRequest request
    ) {

        if (request == null) {
            throw new IllegalArgumentException(
                    "La solicitud es obligatoria"
            );
        }

        if (request.getCodigo() == null
                || request.getCodigo().isBlank()) {

            throw new IllegalArgumentException(
                    "El código es obligatorio"
            );
        }

        if (request.getTipoDocumento() == null
                || request.getTipoDocumento().isBlank()) {

            throw new IllegalArgumentException(
                    "El tipoDocumento es obligatorio"
            );
        }
    }

    private Map<String, String> construirValores(
            EnmiendaExcelData data
    ) {

        Map<String, String> valores =
                new HashMap<>();

        valores.put(
                "Codigo",
                valorSeguro(data.getCodigo())
        );

        valores.put(
                "Titulo",
                valorSeguro(data.getTitulo())
        );

        valores.put(
                "Investigador",
                valorSeguro(data.getInvestigador())
        );

        valores.put(
                "Constancia",
                valorSeguro(data.getConstancia())
        );

        valores.put(
                "fechaCiei",
                valorSeguro(data.getFechaCiei())
        );

        valores.put(
                "fechaIngreso",
                valorSeguro(data.getFechaIngreso())
        );


        return valores;
    }

    private String valorSeguro(String valor) {
        return valor == null ? "" : valor;
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
