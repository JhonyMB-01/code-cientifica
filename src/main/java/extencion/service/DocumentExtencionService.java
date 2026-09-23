package extencion.service;

import contancia.DocumentNotFoundException;
import contancia.config.DocumentConfig;
import extencion.dto.GenerarDocumentoRequest;
import extencion.model.ExtensionExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import org.eclipse.microprofile.config.inject.ConfigProperty;

import java.io.IOException;
import java.nio.file.Path;

@ApplicationScoped
public class DocumentExtencionService {
    @ConfigProperty(name = "documentos.base-path")
    String basePath;

    @ConfigProperty(name = "documentos.excel")
    String excelFile;

    @ConfigProperty(name = "documentos.plantillas-path")
    String plantillasPath;

    @ConfigProperty(name = "documentos.plantilla-extension")
    String plantillaExtension;

    @Inject
    ExtensionExcelService extensionExcelService;

    @Inject
    DocumentConfig documentConfig;

    public byte[] generar(
            GenerarDocumentoRequest request
    ) throws IOException {

        validarRequest(request);

        if (!"EXTENSION".equalsIgnoreCase(
                request.getTipoDocumento())) {

            throw new IllegalArgumentException(
                    "Por el momento solamente se encuentra "
                            + "implementado el documento EXTENSION"
            );
        }

        Path excelPath =
                Path.of(basePath)
                        .resolve(excelFile);

        ExtensionExcelData data =
                extensionExcelService.buscarPorCodigo(
                        excelPath,
                        request.getCodigo()
                );

        if (data == null) {

            throw new DocumentNotFoundException(
                    "No se encontró el código '"
                            + request.getCodigo()
                            + "' en la hoja Extensión"
            );
        }

        /*Path plantilla =
                Path.of(basePath)
                        .resolve(plantillasPath)
                        .resolve(plantillaExtension);*/

        return null;
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

}
