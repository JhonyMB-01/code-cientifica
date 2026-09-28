package informe.controller;

import extencion.dto.GenerarDocumentoRequest;
import extencion.dto.GenerarPdfResponse;
import informe.service.DocumentInformeService;
import jakarta.inject.Inject;
import jakarta.ws.rs.Consumes;
import jakarta.ws.rs.POST;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import utils.GenerarWordResponse;

@Path("/document/informe/v1")
public class InformeController {

    @Inject
    DocumentInformeService documentInformeService;

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

            GenerarWordResponse response = documentInformeService
                    .generarWordInformeAvance(request);

            byte[] documento = response.getPdfContent();
            String nombreArchivo = response.getNameWordGenerate();


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

            GenerarPdfResponse pdfResponse = documentInformeService.generarPdf(request);

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
