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
    @Path("/pdf")
    @Consumes(MediaType.MULTIPART_FORM_DATA)
    @Produces("application/pdf")
    public Response generarPdf(@RestForm("codigo") String codigo,
                               @RestForm("imagenes")List<FileUpload> imagenes) {
        try {

            /*
             * Validar código
             */
            if (codigo == null ||
                    codigo.isBlank()) {

                return Response.status(
                        Response.Status.BAD_REQUEST
                ).entity(
                        "{\"error\":\"El código es obligatorio.\"}"
                ).build();
            }

            /*
             * Si no llegan imágenes,
             * trabajamos con una lista vacía.
             */
            if (imagenes == null) {
                imagenes = new ArrayList<>();
            }

            /*
             * Máximo 6 imágenes.
             */
            if (imagenes.size() > MAX_IMAGENES) {

                return Response.status(
                        Response.Status.BAD_REQUEST
                ).entity(
                        "{\"error\":\"Se permite un máximo de 6 imágenes.\"}"
                ).build();
            }

            /*
             * Convertir FileUpload a DocumentoImagen.
             */
            List<DocumentoImagen> documentosImagen =
                    convertirImagenes(imagenes);

            /*
             * Generar PDF.
             */
            GenerarPdfResponse pdfResponse =
                    enmiendaDocumentService.generarPdf(
                            codigo,
                            documentosImagen
                    );

            byte[] documento = pdfResponse.getPdfContent();
            String nombreArchivo = pdfResponse.getNamePdfGenerate();


            /*
             * Respuesta PDF.
             */
            return Response.ok(documento)
                    .type("application/pdf")
                    .header(
                            "Content-Disposition",
                            "attachment; filename=\"Enmienda_"
                                    + limpiarNombre(codigo)
                                    + ".pdf\""
                    )
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
