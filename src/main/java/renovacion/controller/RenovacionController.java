package renovacion.controller;

import extencion.dto.GenerarDocumentoRequest;
import extencion.dto.GenerarPdfResponse;
import extencion.service.DocumentExtencionService;
import jakarta.inject.Inject;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import renovacion.service.DocumentRenovacionService;

@Path("/document/renovacion/v1")
public class RenovacionController {

    @Inject
    DocumentRenovacionService renovacionService;

    @POST
    @Path("/word")
    @Consumes(MediaType.APPLICATION_JSON)
    @Produces(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    public Response generarWord(
            GenerarDocumentoRequest request) {

        if (request == null || request.getCodigo() == null || request.getCodigo().trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            String codigoLimpio = request.getCodigo().trim();

            byte[] documento =
                    renovacionService.generarWordRenovacion(
                            request
                    );

            String nombreArchivo =
                    "output_" +
                            codigoLimpio +
                            ".docx";

            return Response.ok(documento)
                    .header("Content-Disposition", "attachment; filename=\"" + nombreArchivo + "\"")
                    .build();

        } catch (Exception e) {
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR)
                    .entity("Error al generar el documento: " + e.getMessage())
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

    }

    @POST
    @Path("/pdf")
    @Consumes(MediaType.APPLICATION_JSON)
    @Produces("application/pdf")
    public Response generarPdf(GenerarDocumentoRequest request) {

        if (request == null || request.getCodigo() == null || request.getCodigo().trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            GenerarPdfResponse pdfResponse = renovacionService.generarPdf(request);

            byte[] documento = pdfResponse.getPdfContent();
            String nombreArchivo = pdfResponse.getNamePdfGenerate();

            return Response.ok(documento)
                    .header("Content-Disposition", "attachment; filename=\"" + nombreArchivo + "\"")
                    .build();

        } catch (Exception e) {
            return Response.status(Response.Status.INTERNAL_SERVER_ERROR)
                    .entity("Error al generar el documento: " + e.getMessage())
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

    }
}
