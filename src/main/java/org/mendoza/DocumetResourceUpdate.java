package org.mendoza;

import jakarta.ws.rs.GET;

import jakarta.ws.rs.PathParam;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;
import jakarta.ws.rs.core.Response;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

@jakarta.ws.rs.Path("/document/v2")
public class DocumetResourceUpdate {

    private static final String EXCEL_PATH = "C:/cientifica/Registro_CIEI2024.xlsx";
    private static final String WORD_PATH = "C:/cientifica/Prueba1.docx";
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
                datos.put("Nombre", getCellString(row.getCell(3)));
                datos.put("Descripcion", getCellString(row.getCell(2)));
                return datos;
            }
        }
        return null;
    }

    private File reemplazarWord(Map<String, String> datos, String codigoBuscado) throws Exception {
        try (FileInputStream fis = new FileInputStream(WORD_PATH);
             XWPFDocument doc = new XWPFDocument(fis)) {

            // Reemplazar en párrafos
            for (XWPFParagraph p : doc.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : datos.entrySet()) {
                            String placeholder = "${" + entry.getKey() + "}";
                            if (text.contains(placeholder)) {
                                run.setText(text.replace(placeholder, entry.getValue()), 0);
                            }
                        }
                    }
                }
            }

            // Reemplazar en tablas
            /*for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            for (XWPFRun run : p.getRuns()) {
                                String text = run.getText(0);
                                if (text != null) {
                                    for (Map.Entry<String, String> entry : datos.entrySet()) {
                                        String placeholder = "${" + entry.getKey() + "}";
                                        if (text.contains(placeholder)) {
                                            run.setText(text.replace(placeholder, entry.getValue()), 0);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }*/

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

    private static String getCellString(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: return String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return null;
        }
    }
}
