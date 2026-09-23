package extencion.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExtensionExcelData {

    /**
     * Columna E
     */
    private String codigo;

    /**
     * Columna F
     */
    private String titulo;

    /**
     * Columna G
     */
    private String investigador;

    /**
     * Columna I
     */
    private String constanciaAprobacion;

    /**
     * Columna R
     */
    private String valorR;

    /**
     * Columna T
     */
    private String valorT;

    /**
     * Columna V
     */
    private String valorV;

    /**
     * Columna W
     */
    private String valorW;
}
