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
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;

import java.nio.file.Path;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;


@jakarta.ws.rs.Path("/document")
public class DocumentResource {

    private static final String EXCEL_PATH = "C:/cientifica/Registro_CIEI2024.xlsx";
    //private static final String EXCEL_PATH = "C:/cientifica/Libro1.xlsx";
    private static final String WORD_PATH = "C:/cientifica/documento.docx";
    // Columna F -> índice 5 (A=0)
    private static final int CODIGO_COLUMN = 5;

    @GET
    @jakarta.ws.rs.Path("/{codigo}")
    @Produces(MediaType.APPLICATION_OCTET_STREAM)
    public Response generarPdf(@PathParam("codigo") String codigo) {
        if (codigo == null || codigo.trim().isEmpty()) {
            return Response.status(Response.Status.BAD_REQUEST).entity("Código vacío").build();
        }
        String codigoBuscado = codigo.trim();

        try (FileInputStream fis = new FileInputStream(EXCEL_PATH);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Buscar fila en la columna F (índice 5)
            Map<String, String> datos = null;
            Map<DataFieldName, String> datos2 = null;
            int last = sheet.getLastRowNum();
            for (int r = 7; r <= last; r++) { // empezamos en 1 asumiendo que la fila 0 es cabecera
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(CODIGO_COLUMN);
                String cellValue = getCellString(cell);
                if (cellValue != null && cellValue.equalsIgnoreCase(codigoBuscado)) {
                    /*datos = new HashMap<>();
                    datos.put("Codigo", getCellString(row.getCell(CODIGO_COLUMN)));
                    datos.put("Nombre", getCellString(row.getCell(3)));
                    datos.put("Descripcion", getCellString(row.getCell(2)));*/

                    datos2 = new HashMap<>();
                    datos2.put(new DataFieldName("Codigo"), getCellString(row.getCell(CODIGO_COLUMN)));
                    datos2.put(new DataFieldName("Nombre"), getCellString(row.getCell(3)));
                    datos2.put(new DataFieldName("Descripcion"), getCellString(row.getCell(2)));
                    break;
                }
            }

            if (datos2 == null) {
                return Response.status(Response.Status.NOT_FOUND)
                        .entity("Código no encontrado en Excel").build();
            }

            // Cargar Word
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(WORD_PATH));

            // Reemplazar placeholders en Word
            /*datos.forEach((k, v) -> {
                try {
                    wordMLPackage.getMainDocumentPart().variableReplace(Collections.singletonMap(k, v));
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            });*/

            // Ejecutar Mail Merge
            MailMerger.performMerge(wordMLPackage, datos2, true);

            // Crear carpeta de salida si no existe
            Path outputDir = Path.of("C:/docs");
            if (!Files.exists(outputDir)) {
                Files.createDirectories(outputDir);
            }

            // Guardar en archivo temporal
            File wordFile = new File(outputDir.resolve("output_" + codigoBuscado + ".docx").toString());
            wordMLPackage.save(wordFile);

            // Devolver como descarga
            return Response.ok(wordFile)
                    .header("Content-Disposition", "attachment; filename=" + wordFile.getName())
                    .build();

        } catch (Exception e) {
            return Response.serverError().entity("Error procesando documento: " + e.getMessage()).build();
        }
    }

    // Utilidad para leer celdas de forma segura (maneja nulos y tipos comunes)
    private static String getCellString(Cell cell) {
        if (cell == null) return null;
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return trimToNull(cell.getStringCellValue());
                case NUMERIC:
                    double d = cell.getNumericCellValue();
                    String num;
                    if (d == (long) d) {
                        num = String.valueOf((long) d);
                    } else {
                        num = String.valueOf(d);
                    }
                    return trimToNull(num);
                case BOOLEAN:
                    return trimToNull(String.valueOf(cell.getBooleanCellValue()));
                case FORMULA:
                    try {
                        return trimToNull(cell.getStringCellValue());
                    } catch (Exception ex) {
                        double d2 = cell.getNumericCellValue();
                        return trimToNull(String.valueOf(d2));
                    }
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    private static String trimToNull(String s) {
        if (s == null) return null;
        String t = s.trim();
        return t.isEmpty() ? null : t;
    }

}
