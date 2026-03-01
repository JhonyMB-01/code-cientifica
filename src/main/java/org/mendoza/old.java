package org.mendoza;

import jakarta.ws.rs.GET;
import jakarta.ws.rs.PathParam;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import static org.mendoza.constants.Constantes.*;

@jakarta.ws.rs.Path("/document/v3")
public class old {

    private static final String EXCEL_PATH = "C:/cientifica/Registro_CIEI2024.xlsx";
    private static final String WORD_PATH = "C:/cientifica/documento.docx";
    private static final int CODIGO_COLUMN = 5; // Columna F

    @GET
    @jakarta.ws.rs.Path("/{codigo}")
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
    public Response generarWord(@PathParam("codigo") String codigo) {
        if (codigo == null || codigo.trim().isEmpty()) {
            return Response.status(Response.Status.BAD_REQUEST).entity("Código vacío").build();
        }

        try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Buscar datos en Excel
            Map<String, String> datos = buscarDatos(workbook, codigo.trim());
            if (datos == null) {
                return Response.status(Response.Status.NOT_FOUND).entity("Código no encontrado en Excel").build();
            }

            // Reemplazar placeholders en Word
            File wordFile = reemplazarWord(datos, codigo.trim());

            // Devolver como descarga
            return Response.ok(wordFile)
                    .header("Content-Disposition", "attachment; filename=" + wordFile.getName())
                    .build();

        } catch (Exception e) {
            return Response.serverError().entity("Error procesando documento: " + e.getMessage()).build();
        }
    }

    private Map<String, String> buscarDatos(Workbook workbook, String codigoBuscado) {
        Sheet sheet = workbook.getSheetAt(0);
        int last = sheet.getLastRowNum();

        for (int r = 7; r <= last; r++) { // desde fila 8
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(CODIGO_COLUMN);
            String cellValue = getCellString(cell);
            if (cellValue != null && cellValue.equalsIgnoreCase(codigoBuscado)) {
                Map<String, String> datos = new HashMap<>();
                datos.put("Codigo", getCellString(row.getCell(CODIGO_COLUMN)));
                datos.put("Titulo", getCellString(row.getCell(6)));
                datos.put("Investigador", getCellString(row.getCell(7)));
                datos.put("Constancia", getCellString(row.getCell(33)));

                // Iterar versiones desde v7 hasta v1 y detenerse en la primera no vacía
                //String[] keysVersion = {"7", "6", "5", "4", "3", "2", "1"};
                //int[] cols = {32, 30, 28, 26, 24, 22, 20};

                String[] keysVersion = {"7", "6", "5", "4", "3", "2", "1"};
                int[] cols = {30, 28, 26, 24, 22, 20, 15};

                boolean encontrado = false;
                String ultimaVersion = "";
                for (int i = 0; i < keysVersion.length; i++) {
                    String val = getCellString(row.getCell(cols[i]));
                    if (val != null && !val.isEmpty()) {
                        ultimaVersion = keysVersion[i]+".0 de fecha "+ val;
                        datos.put(keysVersion[i], ultimaVersion);
                        datos.put("ProtocoloInvestigacion", PROT_INVESTIGACION.concat(ultimaVersion));
                        encontrado = true;
                        // Para evitar NPE en el reemplazo, inicializamos las versiones inferiores como cadena vacía
                        for (int j = i + 1; j < keysVersion.length; j++) {
                            datos.put(keysVersion[j], "");
                        }
                        break;
                    } else {
                        // Si no se encontró aún, inicializamos la clave con cadena vacía para evitar NPE
                        datos.put(keysVersion[i], "");
                    }
                }

                // Si ninguna versión tiene valor, asegurarnos de que todas las claves existen (con cadena vacía)
                if (!encontrado) {
                    for (String k : keysVersion) {
                        if (!datos.containsKey(k)) datos.put(k, "");
                    }
                }

                if (CODE_0.equals(getCellString(row.getCell(13)))) {
                    datos.put("ConsentimientoInformado", CONS_INFORMADO.concat(ultimaVersion));
                }else {
                    datos.put("ConsentimientoInformado", "");
                }

                if (CODE_0.equals(getCellString(row.getCell(14)))) {
                    datos.put("AsentimientoInformado", ASEN_INFORMADO.concat(ultimaVersion));
                }else {
                    datos.put("AsentimientoInformado", "");
                }

                String fechaVigencia = getCellString(row.getCell(17));

                datos.put("FechaVigencia", fechaVigencia);
                datos.put("FechaAprobacion", getCellString(row.getCell(16)));

                // Lógica para párrafos de comunicación

                var parrafoPermiso = getCellString(row.getCell(11));
                int orden = 1;

                if("1.1".equals(parrafoPermiso)){
                    datos.put("ParrafoDentroUniversidad", String.valueOf(orden)
                            .concat(SEPARADOR).concat(DENTRO_UNIVERSIDAD));
                    datos.put("ParrafoExternoUniversidad", "");
                    orden++;
                } else if (CODE_0.equals(parrafoPermiso)) {
                    datos.put("ParrafoExternoUniversidad", String.valueOf(orden)
                            .concat(SEPARADOR).concat(EXTERNA_UNIVERSIDAD));
                    datos.put("ParrafoDentroUniversidad", "");
                    orden++;
                } else {
                    datos.put("ParrafoDentroUniversidad", "");
                    datos.put("ParrafoExternoUniversidad", "");
                }

                var parrafoValidacion = getCellString(row.getCell(12));

                if(CODE_0.equals(parrafoValidacion)){
                    datos.put("ParrafoValidacionInstrumento", String.valueOf(orden)
                            .concat(SEPARADOR).concat(VALIDACION_INSTRUMENTOS));
                    orden++;
                } else {
                    datos.put("ParrafoValidacionInstrumento", "");
                }

                datos.put("ParrafoAprobacionEstudio", String.valueOf(orden)
                        .concat(SEPARADOR).concat(APROBACION_ESTUDIO));

                orden++;

                datos.put("ParrafoAprobacionProyecto", String.valueOf(orden)
                        .concat(SEPARADOR).concat(APROBACION_PROYECTO));

                orden++;

                datos.put("ParrafoVigenciaAprobacion", String.valueOf(orden)
                        .concat(SEPARADOR).concat(String.format(VIGENCIA_APROBACION, fechaVigencia)));

                orden++;

                datos.put("ParrafoAprobacionCiei", String.valueOf(orden)
                        .concat(SEPARADOR).concat(APROBACION_CIEI));

                return datos;
            }
        }
        return null;
    }

    private File reemplazarWord(Map<String, String> datos, String codigoBuscado) throws Exception {
        try (FileInputStream fis = new FileInputStream(WORD_PATH);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // Reemplazar en párrafos del cuerpo
            reemplazarEnParrafos(doc.getParagraphs(), datos);

            // Reemplazar en encabezados
            for (XWPFHeader header : doc.getHeaderList()) {
                reemplazarEnParrafos(header.getParagraphs(), datos);
            }

            // Reemplazar en pies de página
            for (XWPFFooter footer : doc.getFooterList()) {
                reemplazarEnParrafos(footer.getParagraphs(), datos);
            }

            // Crear carpeta de salida si no existe
            Path outputDir = Path.of("C:/docs");
            if (!Files.exists(outputDir)) {
                Files.createDirectories(outputDir);
            }

            File wordFile = new File(outputDir.resolve("output_" + codigoBuscado + ".docx").toString());
            try (FileOutputStream fos = new FileOutputStream(wordFile)) {
                doc.write(fos);
            }
            return wordFile;
        }
    }

    private void reemplazarEnParrafos(java.util.List<XWPFParagraph> paragraphs, Map<String, String> datos) {

        for (XWPFParagraph p : paragraphs) {
            String fullText = p.getText();
            if (fullText != null && !fullText.isEmpty()) {
                boolean contienePlaceholder = false;
                for (Map.Entry<String, String> entry : datos.entrySet()) {
                    String placeholder = "${" + entry.getKey() + "}";
                    if (fullText.contains(placeholder)) {
                        fullText = fullText.replace(placeholder, entry.getValue());
                        contienePlaceholder = true;
                    }
                }

                if (contienePlaceholder) {
                    // Guardar estilo del primer run
                    XWPFRun estiloBase = p.getRuns().isEmpty() ? null : p.getRuns().get(0);

                    // Eliminar runs originales
                    int runCount = p.getRuns().size();
                    for (int i = runCount - 1; i >= 0; i--) {
                        p.removeRun(i);
                    }

                    // Crear run nuevo con el texto reemplazado
                    XWPFRun run = p.createRun();
                    run.setText(fullText);

                    // Copiar estilo del run original
                    /*if (estiloBase != null) {
                        run.setBold(estiloBase.isBold());
                        run.setItalic(estiloBase.isItalic());
                        run.setFontFamily(estiloBase.getFontFamily());
                        run.setFontSize(estiloBase.getFontSize());
                        run.setColor(estiloBase.getColor());
                    }*/

                    // Mantener alineación y justificación del párrafo
                    p.setAlignment(p.getAlignment());
                    p.setVerticalAlignment(p.getVerticalAlignment());
                }
            }
        }

    }

    private static String getCellString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    DateFormatSymbols dfs = new DateFormatSymbols(new Locale("es", "ES"));
                    SimpleDateFormat sdf = new SimpleDateFormat("dd 'de' MMMM 'del' yyyy", dfs);
                    return sdf.format(cell.getDateCellValue());
                }
                else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return null;
        }
    }


}
