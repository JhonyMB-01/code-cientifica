package contancia.service;

import contancia.config.DocumentConfig;
import jakarta.enterprise.context.ApplicationScoped;
import jakarta.inject.Inject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.DateFormatSymbols;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

import static org.mendoza.constants.Constantes.*;

@ApplicationScoped
public class ExcelService {

    private static final int CODIGO_COLUMN = 5;

    /*
     * Columnas de versiones.
     */
    private static final String[] KEYS_VERSION = {
            "7", "6", "5", "4", "3", "2", "1"
    };

    private static final int[] VERSION_COLUMNS = {
            30, 28, 26, 24, 22, 20, 15
    };


    private final DataFormatter dataFormatter =
            new DataFormatter(
                    new Locale("es", "PE")
            );

    @Inject
    DocumentConfig documentConfig;


    /**
     * Busca en Excel la información correspondiente
     * al código recibido.
     */
    public Map<String, String> buscarDatos(
            String codigoBuscado) {

        String excelPath =
                documentConfig.getExcelPath();

        try (
                FileInputStream fis =
                        new FileInputStream(excelPath);

                Workbook workbook =
                        new XSSFWorkbook(fis)
        ) {

            return buscarEnWorkbook(
                    workbook,
                    codigoBuscado
            );

        } catch (Exception e) {

            throw new RuntimeException(
                    "No se pudo leer el archivo Excel",
                    e
            );
        }
    }


    private Map<String, String> buscarEnWorkbook(
            Workbook workbook,
            String codigoBuscado) {

        Sheet sheet =
                workbook.getSheetAt(0);

        int last =
                sheet.getLastRowNum();

        /*
         * Desde fila 8.
         * Apache POI utiliza índice 0.
         */
        for (int r = 7; r <= last; r++) {

            Row row =
                    sheet.getRow(r);

            if (row == null) {
                continue;
            }

            Cell cell =
                    row.getCell(CODIGO_COLUMN);

            String cellValue =
                    getCellString(cell);

            if (cellValue == null) {
                continue;
            }

            if (!cellValue.equalsIgnoreCase(
                    codigoBuscado)) {

                continue;
            }

            return construirDatos(row);
        }

        return null;
    }


    private Map<String, String> construirDatos(
            Row row) {

        Map<String, String> datos =
                new HashMap<>();

        /*
         * Datos principales.
         */
        datos.put(
                "Codigo",
                getCellString(
                        row.getCell(CODIGO_COLUMN)
                )
        );

        datos.put(
                "Titulo",
                getCellString(
                        row.getCell(6)
                )
        );

        datos.put(
                "Investigador",
                getCellString(
                        row.getCell(7)
                )
        );

        datos.put(
                "Constancia",
                getCellString(
                        row.getCell(33)
                )
        );


        /*
         * Versiones.
         */
        String ultimaVersion =
                obtenerVersion(row, datos);


        /*
         * Consentimiento informado.
         */
        if (CODE_0.equals(
                getCellString(row.getCell(13)))) {

            datos.put(
                    "ConsentimientoInformado",
                    CONS_INFORMADO.concat(
                            ultimaVersion
                    )
            );

        } else {

            datos.put(
                    "ConsentimientoInformado",
                    ""
            );
        }


        /*
         * Asentimiento informado.
         */
        if (CODE_0.equals(
                getCellString(row.getCell(14)))) {

            datos.put(
                    "AsentimientoInformado",
                    ASEN_INFORMADO.concat(
                            ultimaVersion
                    )
            );

        } else {

            datos.put(
                    "AsentimientoInformado",
                    ""
            );
        }


        /*
         * Fechas.
         */
        String fechaVigencia =
                getCellString(
                        row.getCell(17)
                );

        datos.put(
                "FechaVigencia",
                fechaVigencia
        );

        datos.put(
                "FechaAprobacion",
                getCellString(
                        row.getCell(16)
                )
        );


        /*
         * Párrafos.
         */
        construirParrafos(
                row,
                datos,
                fechaVigencia
        );

        return datos;
    }


    private String obtenerVersion(
            Row row,
            Map<String, String> datos) {

        String ultimaVersion = "";

        boolean encontrada = false;

        for (int i = 0;
             i < KEYS_VERSION.length;
             i++) {

            String key =
                    KEYS_VERSION[i];

            int column =
                    VERSION_COLUMNS[i];

            String valor =
                    getCellString(
                            row.getCell(column)
                    );

            if (valor != null &&
                    !valor.isEmpty()) {

                ultimaVersion =
                        key +
                                ".0 de fecha " +
                                valor;

                datos.put(
                        key,
                        ultimaVersion
                );

                datos.put(
                        "ProtocoloInvestigacion",
                        PROT_INVESTIGACION.concat(
                                ultimaVersion
                        )
                );

                encontrada = true;

                /*
                 * Las versiones inferiores quedan vacías.
                 */
                for (int j = i + 1;
                     j < KEYS_VERSION.length;
                     j++) {

                    datos.put(
                            KEYS_VERSION[j],
                            ""
                    );
                }

                break;

            } else {

                datos.put(
                        key,
                        ""
                );
            }
        }

        /*
         * Si no encontramos ninguna versión.
         */
        if (!encontrada) {

            for (String key : KEYS_VERSION) {

                datos.putIfAbsent(
                        key,
                        ""
                );
            }
        }

        return ultimaVersion;
    }


    private void construirParrafos(
            Row row,
            Map<String, String> datos,
            String fechaVigencia) {

        int orden = 1;


        /*
         * Permiso.
         */
        String parrafoPermiso =
                getCellString(
                        row.getCell(11)
                );

        if ("1.1".equals(parrafoPermiso)) {

            datos.put(
                    "ParrafoDentroUniversidad",
                    orden +
                            SEPARADOR +
                            DENTRO_UNIVERSIDAD
            );

            datos.put(
                    "ParrafoExternoUniversidad",
                    ""
            );

            orden++;

        } else if (
                CODE_0.equals(parrafoPermiso)
        ) {

            datos.put(
                    "ParrafoExternoUniversidad",
                    orden +
                            SEPARADOR +
                            EXTERNA_UNIVERSIDAD
            );

            datos.put(
                    "ParrafoDentroUniversidad",
                    ""
            );

            orden++;

        } else {

            datos.put(
                    "ParrafoDentroUniversidad",
                    ""
            );

            datos.put(
                    "ParrafoExternoUniversidad",
                    ""
            );
        }


        /*
         * Validación del instrumento.
         */
        String parrafoValidacion =
                getCellString(
                        row.getCell(12)
                );

        if (CODE_0.equals(
                parrafoValidacion)) {

            datos.put(
                    "ParrafoValidacionInstrumento",
                    orden +
                            SEPARADOR +
                            VALIDACION_INSTRUMENTOS
            );

            orden++;

        } else {

            datos.put(
                    "ParrafoValidacionInstrumento",
                    ""
            );
        }


        /*
         * Aprobación del estudio.
         */
        datos.put(
                "ParrafoAprobacionEstudio",
                orden +
                        SEPARADOR +
                        APROBACION_ESTUDIO
        );

        orden++;


        /*
         * Aprobación del proyecto.
         */
        datos.put(
                "ParrafoAprobacionProyecto",
                orden +
                        SEPARADOR +
                        APROBACION_PROYECTO
        );

        orden++;


        /*
         * Vigencia.
         */
        datos.put(
                "ParrafoVigenciaAprobacion",
                orden +
                        SEPARADOR +
                        String.format(
                                VIGENCIA_APROBACION,
                                fechaVigencia
                        )
        );

        orden++;


        /*
         * Aprobación CIEI.
         */
        datos.put(
                "ParrafoAprobacionCiei",
                orden +
                        SEPARADOR +
                        APROBACION_CIEI
        );
    }


    /**
     * Convierte una celda Excel a String.
     */
    private String getCellString(Cell cell) {

        if (cell == null) {
            return null;
        }

        if (cell.getCellType() ==
                CellType.BLANK) {

            return "";
        }

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
