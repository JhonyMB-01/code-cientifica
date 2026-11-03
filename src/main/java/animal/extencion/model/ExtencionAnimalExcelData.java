package animal.extencion.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExtencionAnimalExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String ultimaFechaAprobacion;

    private String fechaCiei;

}
