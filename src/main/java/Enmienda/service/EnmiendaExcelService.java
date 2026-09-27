package Enmienda.service;

import Enmienda.model.EnmiendaExcelData;
import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Utils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@ApplicationScoped
public class EnmiendaExcelService {

    private static final String HOJA = "Enmienda";

    public EnmiendaExcelData buscarPorCodigo(Path excelPath, String codigo) throws IOException {

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

                //Obtener fecha CIEI
                String fechaCIEI_R = obtenerValor(row, "R", formatter);
                String fechaCIEI_T = obtenerValor(row, "T", formatter);
                String fechaCIEI_V = obtenerValor(row, "V", formatter);

                String fechaCIEI = "";
                if (!fechaCIEI_V.isBlank()) {
                    fechaCIEI = fechaCIEI_V;
                } else if (!fechaCIEI_T.isBlank()) {
                    fechaCIEI = fechaCIEI_T;
                } else if (!fechaCIEI_R.isBlank()) {
                    fechaCIEI = fechaCIEI_R;
                }

                return  EnmiendaExcelData.builder()
                        .codigo(codigoExcel)
                        .titulo(obtenerValor(row, "F", formatter))
                        .investigador(obtenerValor(row, "G", formatter))
                        .fechaIngreso(Utils.formarterFecha(obtenerValor(row, "I", formatter)))
                        .constancia(obtenerValor(row, "W", formatter))
                        .fechaCiei(Utils.formarterFecha(fechaCIEI))
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