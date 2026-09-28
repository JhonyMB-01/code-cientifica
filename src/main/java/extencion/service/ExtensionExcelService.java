package extencion.service;

import extencion.model.ExtensionExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Utils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@ApplicationScoped
public class ExtensionExcelService {

    private static final String HOJA = "Extensión";

    public ExtensionExcelData buscarPorCodigo(Path excelPath, String codigo) throws IOException {

        if (!Files.exists(excelPath)) {
            throw new IllegalStateException(
                    "No existe el archivo Excel: " + excelPath
            );
        }

        try (InputStream inputStream = Files.newInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheet(HOJA);

            if (sheet == null) {
                throw new IllegalStateException(
                        "No existe la hoja '" + HOJA + "' en el Excel"
                );
            }

            DataFormatter formatter = new DataFormatter();

            for (Row row : sheet) {

                String codigoExcel = obtenerValor(
                        row,
                        "E",
                        formatter
                );

                if (codigoExcel == null || codigoExcel.isBlank()) {
                    continue;
                }

                if (!codigoExcel.trim()
                        .equalsIgnoreCase(codigo.trim())) {

                    continue;
                }

                String fechaCIEI = Utils.primerValorNoVacio(
                        obtenerValor(row, "U", formatter),
                        obtenerValor(row, "S", formatter),
                        obtenerValor(row, "Q", formatter)
                );

                return ExtensionExcelData.builder()
                        .codigo(codigoExcel)
                        .titulo(obtenerValor(row, "F", formatter))
                        .investigador(obtenerValor(row, "G", formatter))
                        .fechaCiei(Utils.formarterFecha(fechaCIEI))
                        .aprHasta(Utils.formarterFecha(obtenerValor(row, "N", formatter)))
                        .constancia(obtenerValor(row, "W", formatter))
                        .build();
            }

            return null;
        }
    }


    private String obtenerValor(
            Row row,
            String columna,
            DataFormatter formatter
    ) {

        int indice = columnaAIndice(columna);

        Cell cell = row.getCell(
                indice,
                Row.MissingCellPolicy.RETURN_BLANK_AS_NULL
        );

        if (cell == null) {
            return "";
        }

        return formatter
                .formatCellValue(cell)
                .trim();
    }

    private int columnaAIndice(String columna) {

        int resultado = 0;

        for (int i = 0; i < columna.length(); i++) {

            resultado = resultado * 26
                    + (columna.charAt(i) - 'A' + 1);
        }

        return resultado - 1;
    }
}