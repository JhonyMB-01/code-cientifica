package animal.aprobacion.service;

import animal.aprobacion.model.AprovacionAnimalExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Utils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@ApplicationScoped
public class AprobacionAnimalExcelService {

    private static final String HOJA = "CIEI-AB";

    public AprovacionAnimalExcelData buscarPorCodigo(Path excelPath, String codigo) throws IOException {

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
                        "G",
                        formatter
                );

                if (codigoExcel.isBlank()) {
                    continue;
                }

                if (!codigoExcel.trim()
                        .equalsIgnoreCase(codigo.trim())) {

                    continue;
                }

                return AprovacionAnimalExcelData.builder()
                        .codigo(codigoExcel)
                        .titulo(obtenerValor(row, "H", formatter))
                        .investigador(obtenerValor(row, "I", formatter))
                        .aprDesde(Utils.formarterFecha(obtenerValor(row, "U", formatter)))
                        .aprHasta(Utils.formarterFecha(obtenerValor(row, "V", formatter)))
                        .constancia(obtenerValor(row, "AL", formatter))
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