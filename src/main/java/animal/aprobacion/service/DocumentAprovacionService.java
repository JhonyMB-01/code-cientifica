package animal.aprobacion.service;

import animal.aprobacion.model.AprovacionAnimalExcelData;
import contancia.DocumentNotFoundException;
import contancia.config.DocumentConfig;
import contancia.service.PdfService;
import contancia.service.WordService;
import extencion.dto.GenerarDocumentoRequest;
import extencion.dto.GenerarPdfResponse;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import renovacion.model.RenovacionExcelData;
import utils.GenerarWordResponse;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import static org.mendoza.constants.Constantes.RENOVACION_PDF_NOMBRE;
import static org.mendoza.constants.Constantes.RENOVACION_WORD_NOMBRE;

@ApplicationScoped
public class DocumentAprovacionService {

    @Inject
    AprobacionAnimalExcelService aprobacionExcelService;

    @Inject
    DocumentConfig documentConfig;

    @Inject
    WordService wordService;

    @Inject
    PdfService pdfService;

    public GenerarWordResponse generarWordRenovacion(GenerarDocumentoRequest request) throws IOException {

        validarRequest(request);

        if (!"RENOVACION".equalsIgnoreCase(
                request.getTipoDocumento())) {

            throw new IllegalArgumentException(
                    "Por el momento solamente se encuentra "
                            + "implementado el documento RENOVACION"
            );
        }

        Path excelPath = Path.of(documentConfig.getBasePath())
                .resolve(documentConfig.getExcelAnimalFile());

        AprovacionAnimalExcelData data =
                aprobacionExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Aprobacion Animal"
            );
        }



        Path plantilla =
                Path.of(documentConfig.getBasePath())
                        .resolve(documentConfig.getPlantillasPath())
                        .resolve(documentConfig.getPlantillaAnimalAprobacion());

        /*
         * 2. Generar DOCX.
         */
        Path docx =
                null;

        try {

            docx =
                    wordService.generarWord(
                            construirValores(data), plantilla.toString(), null
                    );

            return new GenerarWordResponse(Files.readAllBytes(docx),
                    String.format(RENOVACION_WORD_NOMBRE, data.getCodigo(), data.getConstancia()));


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
                .resolve(documentConfig.getExcelAnimalFile());

        AprovacionAnimalExcelData data =
                aprobacionExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Aprobacion Animal"
            );
        }

        Path plantilla =
                Path.of(documentConfig.getBasePath())
                        .resolve(documentConfig.getPlantillasPath())
                        .resolve(documentConfig.getPlantillaAnimalAprobacion());


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
                    String.format(RENOVACION_PDF_NOMBRE, data.getCodigo(), data.getConstancia()));

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

    private Map<String, String> construirValores( AprovacionAnimalExcelData data) {

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
                "AprobHasta",
                valorSeguro(data.getAprHasta())
        );

        valores.put(
                "AprobDesde",
                valorSeguro(data.getAprDesde())
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
