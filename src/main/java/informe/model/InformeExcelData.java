package informe.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class InformeExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String informeAvance;

    private String fechaCiei;

}
