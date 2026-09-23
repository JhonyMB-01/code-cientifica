package contancia.resource;

import contancia.DocumentNotFoundException;
import contancia.service.DocumentService;
import contancia.service.ExcelService;
import contancia.service.WordService;
import extencion.service.DocumentExtencionService;
import jakarta.inject.Inject;
import jakarta.ws.rs.GET;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.PathParam;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;

import java.util.Map;

@Path("/document/v4")
public class DocumentoResource {

    @Inject
    DocumentService documentService;

    /**
     * Genera y descarga el documento Word.
     *
     * GET /document/v3/{codigo}/word
     */
    @GET
    @Path("/{codigo}/word")
    @Produces(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    public Response generarWord(
            @PathParam("codigo") String codigo) {

        if (codigo == null || codigo.trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            String codigoLimpio = codigo.trim();

            byte[] documento =
                    documentService.generarWord(
                            codigoLimpio
                    );

            String nombreArchivo =
                    "output_" +
                            codigoLimpio +
                            ".docx";

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
     * GET /document/v3/{codigo}/pdf
     */
    @GET
    @Path("/{codigo}/pdf")
    @Produces("application/pdf")
    public Response generarPdf(
            @PathParam("codigo") String codigo) {

        if (codigo == null || codigo.trim().isEmpty()) {

            return Response.status(
                            Response.Status.BAD_REQUEST
                    )
                    .entity("Código vacío")
                    .type(MediaType.TEXT_PLAIN)
                    .build();
        }

        try {

            String codigoLimpio =
                    codigo.trim();

            byte[] documento =
                    documentService.generarPdf(
                            codigoLimpio
                    );

            String nombreArchivo =
                    "output_" +
                            codigoLimpio +
                            ".pdf";

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
