package animal.renovacion.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class RenovacionAnimalExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String fechaAproHasta;

    private String FechaUltimaAprobacion;

    private String fechaCiei;

}
