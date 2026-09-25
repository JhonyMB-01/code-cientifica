package renovacion.service;

import extencion.model.ExtensionExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import renovacion.model.RenovacionExcelData;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@ApplicationScoped
public class RenovacionExcelService {

    private static final String HOJA = "Renovación";

    public RenovacionExcelData buscarPorCodigo(Path excelPath, String codigo) throws IOException {

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

                return RenovacionExcelData.builder()
                        .codigo(codigoExcel)
                        .titulo(obtenerValor(row, "F", formatter))
                        .investigador(obtenerValor(row, "G", formatter))
                        .aprDesde(obtenerValor(row, "M", formatter))
                        .aprHasta(obtenerValor(row, "N", formatter))
                        .ventanaDesde(obtenerValor(row, "O", formatter))
                        .ventanaHasta(obtenerValor(row, "P", formatter))
                        .constancia(obtenerValor(row, "X", formatter))
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