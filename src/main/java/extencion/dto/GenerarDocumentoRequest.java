package extencion.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class GenerarDocumentoRequest {

    private String codigo;

    private String tipoDocumento;
}
