package Enmienda.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class EnmiendaExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String fechaIngreso;

    private String fechaCiei;


}
