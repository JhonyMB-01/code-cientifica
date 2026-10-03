package animal.enmienda.service;

import animal.enmienda.model.EnmiendaAnimalExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Utils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@ApplicationScoped
public class EnmiendaAnimalExcelService {

    private static final String HOJA = "Enmienda";

    public EnmiendaAnimalExcelData buscarPorCodigo(Path excelPath, String codigo) throws IOException {

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
                        "F",
                        formatter
                );

                if (codigoExcel.isBlank()) {
                    continue;
                }

                if (!codigoExcel.trim()
                        .equalsIgnoreCase(codigo.trim())) {

                    continue;
                }

                //Obtener fecha CIEI
                String fechaCIEI_R = obtenerValor(row, "S", formatter);
                String fechaCIEI_U = obtenerValor(row, "U", formatter);
                String fechaCIEI_W = obtenerValor(row, "W", formatter);

                String fechaCIEI = "";
                if (!fechaCIEI_W.isBlank()) {
                    fechaCIEI = fechaCIEI_W;
                } else if (!fechaCIEI_U.isBlank()) {
                    fechaCIEI = fechaCIEI_U;
                } else if (!fechaCIEI_R.isBlank()) {
                    fechaCIEI = fechaCIEI_R;
                }

                return EnmiendaAnimalExcelData.builder()
                        .codigo(codigoExcel)
                        .titulo(obtenerValor(row, "G", formatter))
                        .investigador(obtenerValor(row, "H", formatter))
                        .constancia(obtenerValor(row, "X", formatter))
                        .fechaCIEI(Utils.formarterFecha(fechaCIEI))
                        .fechaIngreso(Utils.formarterFecha(obtenerValor(row, "J", formatter)))
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