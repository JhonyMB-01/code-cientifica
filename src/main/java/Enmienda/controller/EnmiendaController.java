package Enmienda.controller;

import Enmienda.model.DocumentoImagen;
import Enmienda.service.EnmiendaDocumentService;
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
import org.jboss.resteasy.reactive.RestForm;
import org.jboss.resteasy.reactive.multipart.FileUpload;
import utils.GenerarWordResponse;

import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

@Path("/document/enmienda/v1")
public class EnmiendaController {

    @Inject
    EnmiendaDocumentService enmiendaDocumentService;

    private static final int MAX_IMAGENES = 6;

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

            GenerarWordResponse response = enmiendaDocumentService.generarWordEnmienda(request);
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

            /*
             * Generar PDF.
             */
            GenerarPdfResponse pdfResponse =
                    enmiendaDocumentService.generarPdf(request);

            byte[] documento = pdfResponse.getPdfContent();
            String nombreArchivo = pdfResponse.getNamePdfGenerate();


            /*
             * Respuesta PDF.
             */
            return Response.ok(documento)
                    .header("Content-Disposition", "attachment; filename=\"" + nombreArchivo + "\"")
                    .build();

        } catch (IllegalArgumentException e) {

            return Response.status(
                    Response.Status.BAD_REQUEST
            ).entity(
                    "{\"error\":\""
                            + escaparJson(e.getMessage())
                            + "\"}"
            ).build();

        } catch (IOException e) {

            return Response.status(
                    Response.Status.INTERNAL_SERVER_ERROR
            ).entity(
                    "{\"error\":\"No se pudo procesar el documento.\"}"
            ).build();

        } catch (Exception e) {

            return Response.status(
                    Response.Status.INTERNAL_SERVER_ERROR
            ).entity(
                    "{\"error\":\"Ocurrió un error al generar la enmienda.\"}"
            ).build();
        }
    }

    private List<DocumentoImagen> convertirImagenes(
            List<FileUpload> uploads)
            throws IOException {

        List<DocumentoImagen> resultado =
                new ArrayList<>();

        for (FileUpload upload : uploads) {

            if (upload == null) {
                continue;
            }

            DocumentoImagen imagen =
                    DocumentoImagen.builder()
                            .nombre(
                                    upload.fileName()
                            )
                            .contentType(
                                    upload.contentType()
                            )
                            .contenido(
                                    Files.readAllBytes(
                                            upload.uploadedFile()
                                    )
                            )
                            .build();

            resultado.add(imagen);
        }

        return resultado;
    }

    private String limpiarNombre(
            String valor) {

        if (valor == null ||
                valor.isBlank()) {

            return "documento";
        }

        return valor.replaceAll(
                "[^a-zA-Z0-9._-]",
                "_"
        );
    }

    private String escaparJson(
            String valor) {

        if (valor == null) {
            return "";
        }

        return valor
                .replace("\\", "\\\\")
                .replace("\"", "\\\"");
    }

}
