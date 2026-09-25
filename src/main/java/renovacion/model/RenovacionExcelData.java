package renovacion.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class RenovacionExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String aprHasta;

    private String aprDesde;

    private String ventanaDesde;

    private String ventanaHasta;
}
