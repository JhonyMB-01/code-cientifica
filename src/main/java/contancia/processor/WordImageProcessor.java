package contancia.processor;

import Enmienda.model.DocumentoImagen;

import jakarta.enterprise.context.ApplicationScoped;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import com.drew.imaging.ImageMetadataReader;
import com.drew.metadata.Metadata;
import com.drew.metadata.exif.ExifIFD0Directory;
import javax.imageio.ImageIO;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

import java.util.List;

@ApplicationScoped
public class WordImageProcessor {

    /*
     * ============================================================
     * CONFIGURACIÓN
     * ============================================================
     *
     * Tamaño máximo de una imagen:
     *
     * Ancho máximo: 15 cm
     * Alto máximo: 10 cm
     *
     * IMPORTANTE:
     * Estos límites NO rotan la fotografía.
     * Solamente limitan su tamaño conservando proporción.
     */
    private static final int MAX_ANCHO_EMU =
            Units.toEMU(15 * 28.3464567);

    private static final int MAX_ALTO_EMU =
            Units.toEMU(10 * 28.3464567);


    /*
     * ============================================================
     * MÉTODO PRINCIPAL
     * ============================================================
     */
    public void procesarImagenes(
            XWPFDocument document,
            List<DocumentoImagen> imagenes) {

        if (document == null) {
            return;
        }

        /*
         * --------------------------------------------------------
         * 1. PÁRRAFOS NORMALES
         * --------------------------------------------------------
         */
        for (int i = 1; i <= 6; i++) {

            DocumentoImagen imagen =
                    obtenerImagen(
                            imagenes,
                            i
                    );

            String placeholder =
                    "${Imagen" + i + "}";

            procesarPlaceholderEnParrafos(
                    document.getParagraphs(),
                    placeholder,
                    imagen
            );
        }


        /*
         * --------------------------------------------------------
         * 2. TABLAS
         * --------------------------------------------------------
         *
         * Si algún ${ImagenX} está dentro de una tabla,
         * document.getParagraphs() no lo encuentra.
         */
        for (XWPFTable table :
                document.getTables()) {

            procesarTabla(
                    table,
                    imagenes
            );
        }
    }


    /*
     * ============================================================
     * PROCESAR TABLAS
     * ============================================================
     */
    private void procesarTabla(
            XWPFTable table,
            List<DocumentoImagen> imagenes) {

        if (table == null) {
            return;
        }

        for (XWPFTableRow row :
                table.getRows()) {

            for (XWPFTableCell cell :
                    row.getTableCells()) {

                /*
                 * Procesar los párrafos de la celda.
                 */
                for (int i = 1; i <= 6; i++) {

                    DocumentoImagen imagen =
                            obtenerImagen(
                                    imagenes,
                                    i
                            );

                    String placeholder =
                            "${Imagen" + i + "}";

                    procesarPlaceholderEnParrafos(
                            cell.getParagraphs(),
                            placeholder,
                            imagen
                    );
                }


                /*
                 * Procesar tablas anidadas.
                 */
                for (XWPFTable tablaInterna :
                        cell.getTables()) {

                    procesarTabla(
                            tablaInterna,
                            imagenes
                    );
                }
            }
        }
    }


    /*
     * ============================================================
     * OBTENER IMAGEN
     * ============================================================
     */
    private DocumentoImagen obtenerImagen(
            List<DocumentoImagen> imagenes,
            int posicion) {

        if (imagenes == null) {
            return null;
        }

        if (imagenes.size() < posicion) {
            return null;
        }

        return imagenes.get(posicion - 1);
    }


    /*
     * ============================================================
     * PROCESAR PLACEHOLDER
     * ============================================================
     */
    private void procesarPlaceholderEnParrafos(
            List<XWPFParagraph> paragraphs,
            String placeholder,
            DocumentoImagen imagen) {

        if (paragraphs == null) {
            return;
        }

        for (XWPFParagraph paragraph :
                paragraphs) {

            if (!contienePlaceholder(
                    paragraph,
                    placeholder)) {

                continue;
            }


            /*
             * Si no existe imagen para ese placeholder,
             * simplemente eliminamos ${ImagenX}.
             */
            if (imagen == null ||
                    imagen.getContenido() == null ||
                    imagen.getContenido().length == 0) {

                limpiarParrafo(
                        paragraph
                );

                return;
            }


            try {

                insertarImagen(
                        paragraph,
                        imagen
                );

            } catch (Exception e) {

                throw new RuntimeException(
                        "No se pudo insertar la imagen "
                                + placeholder
                                + ". Archivo: "
                                + imagen.getNombre(),
                        e
                );
            }

            return;
        }
    }


    /*
     * ============================================================
     * BUSCAR PLACEHOLDER
     * ============================================================
     */
    private boolean contienePlaceholder(
            XWPFParagraph paragraph,
            String placeholder) {

        if (paragraph == null) {
            return false;
        }

        StringBuilder texto =
                new StringBuilder();

        for (XWPFRun run :
                paragraph.getRuns()) {

            if (run == null) {
                continue;
            }

            String textoRun =
                    obtenerTextoRun(run);

            if (textoRun != null) {
                texto.append(textoRun);
            }
        }

        return texto
                .toString()
                .contains(placeholder);
    }


    /*
     * ============================================================
     * INSERTAR IMAGEN
     * ============================================================
     */
    private void insertarImagen(
            XWPFParagraph paragraph,
            DocumentoImagen imagen)
            throws Exception {

        byte[] contenidoOriginal =
                imagen.getContenido();


        /*
         * --------------------------------------------------------
         * 1. CORREGIR ORIENTACIÓN EXIF
         * --------------------------------------------------------
         *
         * Esta es la parte importante para fotografías
         * tomadas con celulares.
         */
        BufferedImage imagenCorregida =
                corregirOrientacion(
                        contenidoOriginal
                );


        /*
         * --------------------------------------------------------
         * 2. DETERMINAR FORMATO
         * --------------------------------------------------------
         */
        String contentType =
                normalizarContentType(
                        imagen.getContentType()
                );


        /*
         * --------------------------------------------------------
         * 3. CONVERTIR LA IMAGEN CORREGIDA
         *    NUEVAMENTE A BYTES
         * --------------------------------------------------------
         *
         * No podemos insertar los bytes originales porque
         * volveríamos a tener el problema de EXIF.
         */
        byte[] contenidoCorregido =
                convertirImagenABytes(
                        imagenCorregida,
                        contentType
                );


        /*
         * --------------------------------------------------------
         * 4. DIMENSIONES REALES
         * --------------------------------------------------------
         */
        int anchoOriginal =
                imagenCorregida.getWidth();

        int altoOriginal =
                imagenCorregida.getHeight();


        /*
         * --------------------------------------------------------
         * 5. CALCULAR DIMENSIONES SIN DEFORMAR
         * --------------------------------------------------------
         */
        int[] dimensiones =
                calcularDimensiones(
                        anchoOriginal,
                        altoOriginal
                );


        /*
         * --------------------------------------------------------
         * 6. LIMPIAR PLACEHOLDER
         * --------------------------------------------------------
         */
        limpiarParrafo(
                paragraph
        );


        /*
         * --------------------------------------------------------
         * 7. CREAR RUN
         * --------------------------------------------------------
         */
        XWPFRun run =
                paragraph.createRun();


        /*
         * --------------------------------------------------------
         * 8. INSERTAR IMAGEN
         * --------------------------------------------------------
         */
        run.addPicture(
                new ByteArrayInputStream(
                        contenidoCorregido
                ),
                obtenerTipoImagen(
                        contentType
                ),
                obtenerNombreImagen(
                        imagen.getNombre(),
                        contentType
                ),
                dimensiones[0],
                dimensiones[1]
        );
    }


    /*
     * ============================================================
     * CORREGIR ORIENTACIÓN EXIF
     * ============================================================
     */
    private BufferedImage corregirOrientacion(
            byte[] contenido) throws Exception {

        BufferedImage original =
                ImageIO.read(
                        new ByteArrayInputStream(
                                contenido
                        )
                );

        if (original == null) {

            throw new IllegalArgumentException(
                    "El archivo no es una imagen válida."
            );
        }


        /*
         * Por defecto:
         *
         * Orientation = 1
         */
        int orientacion = 1;


        try {

            Metadata metadata =
                    ImageMetadataReader.readMetadata(
                            new ByteArrayInputStream(
                                    contenido
                            )
                    );

            ExifIFD0Directory exif =
                    metadata.getFirstDirectoryOfType(
                            ExifIFD0Directory.class
                    );

            if (exif != null &&
                    exif.containsTag(
                            ExifIFD0Directory.TAG_ORIENTATION
                    )) {

                orientacion =
                        exif.getInt(
                                ExifIFD0Directory.TAG_ORIENTATION
                        );
            }

        } catch (Exception ignored) {

            /*
             * Si la imagen no tiene EXIF o el EXIF
             * no puede leerse, utilizamos la imagen
             * original.
             */
            orientacion = 1;
        }


        /*
         * No necesita transformación.
         */
        if (orientacion == 1) {
            return original;
        }


        return aplicarOrientacion(
                original,
                orientacion
        );
    }


    /*
     * ============================================================
     * APLICAR ORIENTACIÓN EXIF
     * ============================================================
     */
    private BufferedImage aplicarOrientacion(
            BufferedImage original,
            int orientacion) {

        int ancho =
                original.getWidth();

        int alto =
                original.getHeight();


        /*
         * Las orientaciones 5, 6, 7 y 8 intercambian
         * ancho por alto.
         */
        boolean intercambiarDimensiones =
                orientacion >= 5 &&
                        orientacion <= 8;


        int nuevoAncho =
                intercambiarDimensiones
                        ? alto
                        : ancho;

        int nuevoAlto =
                intercambiarDimensiones
                        ? ancho
                        : alto;


        BufferedImage resultado =
                new BufferedImage(
                        nuevoAncho,
                        nuevoAlto,
                        BufferedImage.TYPE_INT_RGB
                );


        Graphics2D graphics =
                resultado.createGraphics();


        graphics.setRenderingHint(
                RenderingHints.KEY_INTERPOLATION,
                RenderingHints.VALUE_INTERPOLATION_BILINEAR
        );

        graphics.setRenderingHint(
                RenderingHints.KEY_RENDERING,
                RenderingHints.VALUE_RENDER_QUALITY
        );

        graphics.setRenderingHint(
                RenderingHints.KEY_ANTIALIASING,
                RenderingHints.VALUE_ANTIALIAS_ON
        );


        AffineTransform transform =
                new AffineTransform();


        switch (orientacion) {

            /*
             * ----------------------------------------------------
             * 2 = Mirror horizontal
             * ----------------------------------------------------
             */
            case 2:

                transform.scale(
                        -1,
                        1
                );

                transform.translate(
                        -ancho,
                        0
                );

                break;


            /*
             * ----------------------------------------------------
             * 3 = Rotate 180°
             * ----------------------------------------------------
             */
            case 3:

                transform.rotate(
                        Math.PI
                );

                transform.translate(
                        -ancho,
                        -alto
                );

                break;


            /*
             * ----------------------------------------------------
             * 4 = Mirror vertical
             * ----------------------------------------------------
             */
            case 4:

                transform.scale(
                        1,
                        -1
                );

                transform.translate(
                        0,
                        -alto
                );

                break;


            /*
             * ----------------------------------------------------
             * 5 = Transpose
             * ----------------------------------------------------
             */
            case 5:

                transform.rotate(
                        Math.PI / 2
                );

                transform.scale(
                        -1,
                        1
                );

                break;


            /*
             * ----------------------------------------------------
             * 6 = Rotate 90° CW
             * ----------------------------------------------------
             */
            case 6:

                transform.rotate(
                        Math.PI / 2
                );

                transform.translate(
                        0,
                        -alto
                );

                break;


            /*
             * ----------------------------------------------------
             * 7 = Transverse
             * ----------------------------------------------------
             */
            case 7:

                transform.rotate(
                        -Math.PI / 2
                );

                transform.scale(
                        -1,
                        1
                );

                break;


            /*
             * ----------------------------------------------------
             * 8 = Rotate 270° CW
             * ----------------------------------------------------
             */
            case 8:

                transform.rotate(
                        -Math.PI / 2
                );

                transform.translate(
                        -ancho,
                        0
                );

                break;


            /*
             * ----------------------------------------------------
             * Desconocido
             * ----------------------------------------------------
             */
            default:

                graphics.dispose();

                return original;
        }


        graphics.drawImage(
                original,
                transform,
                null
        );


        graphics.dispose();


        return resultado;
    }


    /*
     * ============================================================
     * CONVERTIR IMAGEN A BYTES
     * ============================================================
     */
    private byte[] convertirImagenABytes(
            BufferedImage imagen,
            String contentType)
            throws IOException {

        String formato;


        if ("image/png".equals(contentType)) {

            formato = "png";

        } else {

            /*
             * JPEG será el formato por defecto.
             */
            formato = "jpg";
        }


        /*
         * JPEG no soporta transparencia.
         *
         * Creamos una imagen RGB con fondo blanco.
         */
        if ("jpg".equals(formato)) {

            BufferedImage rgb =
                    new BufferedImage(
                            imagen.getWidth(),
                            imagen.getHeight(),
                            BufferedImage.TYPE_INT_RGB
                    );

            Graphics2D graphics =
                    rgb.createGraphics();

            graphics.setColor(
                    java.awt.Color.WHITE
            );

            graphics.fillRect(
                    0,
                    0,
                    rgb.getWidth(),
                    rgb.getHeight()
            );

            graphics.drawImage(
                    imagen,
                    0,
                    0,
                    null
            );

            graphics.dispose();

            imagen = rgb;
        }


        ByteArrayOutputStream output =
                new ByteArrayOutputStream();

        boolean escrito =
                ImageIO.write(
                        imagen,
                        formato,
                        output
                );

        if (!escrito) {

            throw new IOException(
                    "No se pudo convertir la imagen a "
                            + formato
            );
        }


        return output.toByteArray();
    }


    /*
     * ============================================================
     * CALCULAR DIMENSIONES
     * ============================================================
     *
     * Mantiene SIEMPRE la relación de aspecto.
     *
     * Ejemplo:
     *
     * Foto vertical:
     *     3024 x 4032
     *
     * Foto horizontal:
     *     4032 x 3024
     *
     * Ninguna será deformada.
     */
    private int[] calcularDimensiones(
            int anchoOriginal,
            int altoOriginal) {

        if (anchoOriginal <= 0 ||
                altoOriginal <= 0) {

            throw new IllegalArgumentException(
                    "Las dimensiones de la imagen no son válidas."
            );
        }


        double escalaAncho =
                (double) MAX_ANCHO_EMU
                        / Units.toEMU(
                        anchoOriginal
                );


        double escalaAlto =
                (double) MAX_ALTO_EMU
                        / Units.toEMU(
                        altoOriginal
                );


        /*
         * Usamos la escala más pequeña para garantizar
         * que la imagen entre completamente dentro
         * del área máxima.
         */
        double escala =
                Math.min(
                        escalaAncho,
                        escalaAlto
                );


        /*
         * No agrandar fotografías pequeñas.
         */
        escala =
                Math.min(
                        escala,
                        1.0
                );


        int ancho =
                (int) (
                        Units.toEMU(
                                anchoOriginal
                        ) * escala
                );


        int alto =
                (int) (
                        Units.toEMU(
                                altoOriginal
                        ) * escala
                );


        return new int[]{
                ancho,
                alto
        };
    }


    /*
     * ============================================================
     * CONTENT-TYPE
     * ============================================================
     */
    private String normalizarContentType(
            String contentType) {

        if (contentType == null ||
                contentType.isBlank()) {

            /*
             * Si no viene Content-Type, asumimos JPEG.
             */
            return "image/jpeg";
        }


        return contentType
                .trim()
                .toLowerCase();
    }


    /*
     * ============================================================
     * TIPO DE IMAGEN PARA APACHE POI
     * ============================================================
     */
    private int obtenerTipoImagen(
            String contentType) {

        return switch (contentType) {

            case "image/png" ->
                    XWPFDocument.PICTURE_TYPE_PNG;

            case "image/jpeg",
                 "image/jpg" ->
                    XWPFDocument.PICTURE_TYPE_JPEG;

            default ->
                    throw new IllegalArgumentException(
                            "Formato de imagen no soportado: "
                                    + contentType
                    );
        };
    }


    /*
     * ============================================================
     * NOMBRE DE IMAGEN
     * ============================================================
     */
    private String obtenerNombreImagen(
            String nombre,
            String contentType) {

        if (nombre != null &&
                !nombre.isBlank()) {

            return nombre;
        }


        if ("image/png".equals(contentType)) {
            return "imagen.png";
        }


        return "imagen.jpg";
    }


    /*
     * ============================================================
     * LIMPIAR PÁRRAFO
     * ============================================================
     */
    private void limpiarParrafo(
            XWPFParagraph paragraph) {

        if (paragraph == null) {
            return;
        }


        /*
         * Eliminamos los runs completos.
         *
         * Esto es más seguro que:
         *
         * run.setText("", 0)
         *
         * porque evita dejar residuos del placeholder.
         */
        for (int i =
             paragraph.getRuns().size() - 1;
             i >= 0;
             i--) {

            paragraph.removeRun(i);
        }
    }


    /*
     * ============================================================
     * OBTENER TEXTO DE RUN
     * ============================================================
     */
    private String obtenerTextoRun(
            XWPFRun run) {

        if (run == null) {
            return "";
        }


        StringBuilder texto =
                new StringBuilder();


        int cantidad =
                run.getCTR()
                        .sizeOfTArray();


        for (int i = 0;
             i < cantidad;
             i++) {

            String valor =
                    run.getText(i);


            if (valor != null) {
                texto.append(valor);
            }
        }


        return texto.toString();
    }
}
