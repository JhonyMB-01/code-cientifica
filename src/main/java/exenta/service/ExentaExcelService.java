package exenta.service;

import jakarta.enterprise.context.ApplicationScoped;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import utils.Utils;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

import static org.apache.logging.log4j.util.Strings.EMPTY;
import static org.mendoza.constants.Constantes.*;
import static org.mendoza.constants.Constantes.APROBACION_CIEI;
import static org.mendoza.constants.Constantes.APROBACION_ESTUDIO;
import static org.mendoza.constants.Constantes.APROBACION_PROYECTO;
import static org.mendoza.constants.Constantes.CODE_0;
import static org.mendoza.constants.Constantes.EXTERNA_UNIVERSIDAD;
import static org.mendoza.constants.Constantes.SEPARADOR;
import static org.mendoza.constants.Constantes.VALIDACION_INSTRUMENTOS;
import static org.mendoza.constants.Constantes.VIGENCIA_APROBACION;

@ApplicationScoped
public class ExentaExcelService {

    private static final String HOJA = "1.-Registro 22-23";

    /*
     * Columnas de versiones.
     */
    private static final String[] KEYS_VERSION = {
            "6", "5", "4", "3", "2", "1"
    };

    private static final String[] VERSION_COLUMNS = {
            "AL", "AJ", "AH", "AF", "AD", "AB"
    };

    public Map<String, String> buscarPorCodigo(Path excelPath, String codigo) throws IOException {

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

                if (codigoExcel == null || codigoExcel.isBlank()) {
                    continue;
                }

                if (!codigoExcel.trim()
                        .equalsIgnoreCase(codigo.trim())) {

                    continue;
                }

                return construirDatos(row, formatter);

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

    private Map<String, String> construirDatos(
            Row row, DataFormatter formatter) {

        Map<String, String> datos =  new HashMap<>();

        /*
         * Datos principales.
         */
        datos.put(
                "Codigo",  obtenerValor(row, "G", formatter)
        );

        datos.put(
                "Titulo", obtenerValor(row, "H", formatter)
        );

        datos.put(
                "Investigador", obtenerValor(row, "I", formatter)
        );

        datos.put(
                "Constancia", obtenerValor(row, "AO", formatter)
        );


        /*
         * Versiones.
         */
        String ultimaVersion =
                obtenerVersion(row, datos, formatter);


        /*
         * Consentimiento informado.
         */
        if (CODE_0.equals(obtenerValor(row, "Q", formatter))) {

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
        if (CODE_0.equals(obtenerValor(row, "R", formatter))) {

            datos.put(
                    "AsentimientoInformado", ASEN_INFORMADO.concat(ultimaVersion)
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
        String fechaVigencia = obtenerValor(row, "Z", formatter);

        datos.put(
                "FechaVigencia", Utils.formarterFecha(fechaVigencia)
        );

        datos.put(
                "FechaAprobacion",
                Utils.formarterFecha(obtenerValor(row, "Y", formatter))
        );


        /*
         * Párrafos.
         */
        construirParrafos(
                row,
                datos,
                fechaVigencia,
                formatter
        );

        return datos;
    }

    private String obtenerVersion(
            Row row, Map<String, String> datos,  DataFormatter formatter) {

        String ultimaVersion = "";

        boolean encontrada = false;

        for (int i = 0;
             i < KEYS_VERSION.length;
             i++) {

            String key =
                    KEYS_VERSION[i];

            String column =
                    VERSION_COLUMNS[i];

            String valor = obtenerValor(row, column, formatter);

            if (!valor.isEmpty()) {

                ultimaVersion =
                        key +
                                ".0 de fecha " +
                                Utils.formarterFecha(valor);

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
            String fechaVigencia, DataFormatter formatter) {

        int orden = 1;


        /*
         * Permiso.
         */
        String parrafoPermiso = obtenerValor(row, "M", formatter);


        if ("1".equals(parrafoPermiso)) {

            datos.put(
                    "ParrafoDentroUniversidad",
                    orden +
                            SEPARADOR +
                            DENTRO_UNIVERSIDAD
            );

            datos.put(
                    "ParrafoExternoUniversidad",
                    EMPTY.strip()
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
                    EMPTY.strip()
            );

            orden++;

        } else {

            datos.put(
                    "ParrafoDentroUniversidad",
                    EMPTY.strip()
            );

            datos.put(
                    "ParrafoExternoUniversidad",
                    EMPTY.strip()
            );
        }


        /*
         * Validación del instrumento.
         */
        String parrafoValidacion =
                obtenerValor(row, "N", formatter);

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
                    "".strip()
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
                                Utils.formarterFecha(fechaVigencia)
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

}