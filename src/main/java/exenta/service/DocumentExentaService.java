package exenta.service;

import contancia.DocumentNotFoundException;
import contancia.config.DocumentConfig;
import contancia.service.PdfService;
import contancia.service.WordService;
import extencion.dto.GenerarDocumentoRequest;
import extencion.dto.GenerarPdfResponse;
import extencion.model.ExtensionExcelData;
import extencion.service.ExtensionExcelService;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import utils.GenerarWordResponse;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import static org.mendoza.constants.Constantes.*;

@ApplicationScoped
public class DocumentExentaService {

    @Inject
    ExentaExcelService exentaExcelService;

    @Inject
    DocumentConfig documentConfig;

    @Inject
    WordService wordService;

    @Inject
    PdfService pdfService;

    public GenerarWordResponse generarWord(GenerarDocumentoRequest request) throws IOException {

        validarRequest(request);

        /*if (!"EXENTA".equalsIgnoreCase(
                request.getTipoDocumento())) {

            throw new IllegalArgumentException(
                    "Por el momento solamente se encuentra "
                            + "implementado el documento EXENTA"
            );
        }*/

        Path excelPath = Path.of(documentConfig.getBasePath())
                .resolve(documentConfig.getExcelFile());

        Map<String, String> data =
                exentaExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Exenta"
            );
        }

        Path plantilla = getPath(request);


        /*
         * 2. Generar DOCX.
         */
        Path docx =
                null;

        try {

            docx =
                    wordService.generarWord(
                            data, plantilla.toString(), null
                    );

            String wordName = request.getTipoDocumento().equals(EXENTA_NOMBRE)
                    ? String.format(EXENTA_WORD_NOMBRE, data.get("Codigo"), data.get("Constancia"))
                    : String.format(CONSTANCIA_WORD_NOMBRE, data.get("Codigo"), data.get("Constancia"));

            return new GenerarWordResponse(Files.readAllBytes(docx),
                    wordName);

        } finally {

            eliminarTemporal(
                    docx
            );
        }
    }

    private Path getPath(GenerarDocumentoRequest request) {
        Path plantilla;

        if (request.getTipoDocumento().equalsIgnoreCase(EXENTA_NOMBRE)) {
            plantilla =
                    Path.of(documentConfig.getBasePath())
                            .resolve(documentConfig.getPlantillasPath())
                            .resolve(documentConfig.getPlantillaConstanciaExenta());
        }else {
            plantilla =
                    Path.of(documentConfig.getBasePath())
                            .resolve(documentConfig.getPlantillasPath())
                            .resolve(documentConfig.getPlantillaConstancia());
        }
        return plantilla;
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

        Map<String, String> data =
                exentaExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Exenta"
            );
        }

        Path plantilla = getPath(request);


        Path docx = null;

        try {

            /*
             * 2. Generar DOCX.
             */
            docx =
                    wordService.generarWord(
                            data, plantilla.toString(), null
                    );

            String pdfName = request.getTipoDocumento().equalsIgnoreCase(EXENTA_NOMBRE)
                    ? String.format(EXENTA_PDF_NOMBRE, data.get("Codigo"), data.get("Constancia"))
                    : String.format(CONSTANCIA_PDF_NOMBRE, data.get("Codigo"), data.get("Constancia"));


            /*
             * 3. Convertir DOCX → PDF.
             */
            return new GenerarPdfResponse(
                    pdfService.convertirDocxAPdf(docx),
                    pdfName);

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
