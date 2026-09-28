package contancia.resource;

import contancia.DocumentNotFoundException;
import contancia.service.DocumentService;
import exenta.service.DocumentExentaService;
import extencion.dto.GenerarDocumentoRequest;

import extencion.dto.GenerarPdfResponse;
import jakarta.inject.Inject;
import jakarta.ws.rs.*;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import utils.GenerarWordResponse;

import java.util.Map;

@Path("/document/constancia/v1")
public class DocumentoResource {

    @Inject
    DocumentExentaService exentaService;

    /**
     * Genera y descarga el documento Word.
     *
     * POST /document/v3/word
     */
    @POST
    @Path("/word")
    @Consumes(MediaType.APPLICATION_JSON)
    @Produces(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    public Response generarWord(GenerarDocumentoRequest request) {

        String codigo = request.getCodigo();

        if (codigo == null || codigo.trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            GenerarWordResponse response =
                    exentaService.generarWord(
                            request
                    );

            byte[] documento = response.getPdfContent();
            String nombreArchivo = response.getNameWordGenerate();

            return Response.ok(documento)
                    .type(
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    .header(
                            "Content-Disposition",
                            "attachment; filename=\"" +
                                    nombreArchivo +
                                    "\""
                    )
                    .build();

        } catch (DocumentNotFoundException e) {

            return Response.status(
                            Response.Status.NOT_FOUND
                    )
                    .entity(e.getMessage())
                    .type(MediaType.TEXT_PLAIN)
                    .build();

        } catch (Exception e) {

            return Response.status(
                            Response.Status.INTERNAL_SERVER_ERROR
                    )
                    .entity(
                            "Error generando Word: " +
                                    e.getMessage()
                    )
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }
    }


    /**
     * Genera y descarga el documento PDF.
     *
     * POST /document/v3/pdf
     */
    @POST
    @Path("/pdf")
    @Consumes(MediaType.APPLICATION_JSON)
    @Produces("application/pdf")
    public Response generarPdf(GenerarDocumentoRequest request) {

        String codigo = request.getCodigo();

        if (codigo == null || codigo.trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            GenerarPdfResponse response = exentaService.generarPdf(request);

            byte[] documento = response.getPdfContent();
            String nombreArchivo = response.getNamePdfGenerate();

            return Response.ok(documento)
                    .type("application/pdf")
                    .header(
                            "Content-Disposition",
                            "attachment; filename=\"" +
                                    nombreArchivo +
                                    "\""
                    )
                    .build();

        } catch (DocumentNotFoundException e) {

            return Response.status(
                            Response.Status.NOT_FOUND
                    )
                    .entity(e.getMessage())
                    .type(MediaType.TEXT_PLAIN)
                    .build();

        } catch (Exception e) {

            return Response.status(
                            Response.Status.INTERNAL_SERVER_ERROR
                    )
                    .entity(
                            "Error generando PDF: " +
                                    e.getMessage()
                    )
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }
    }
}
