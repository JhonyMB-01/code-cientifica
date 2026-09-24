package extencion.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class GenerarPdfResponse {

    byte[] pdfContent;
    String namePdfGenerate;
}
