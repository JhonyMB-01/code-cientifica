package exenta.model;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExentanExcelData {

    private String codigo;

    private String titulo;

    private String investigador;

    private String constancia;

    private String protocoloInvestigacion;

    private String consentimientoInformado;

    private String asentamientoInformado;

    private String fechaAprobacionHasta;

    private String fechaAprobacionDesde;

    private String parrafoDentroUni;

    private String parrafoExternoUni;

    private String parrafoValidacionInstrumento;

    private String parrafoAprobacionEstudio;

    private String parrafoVigenciaAProb;

    private String parrafoAprobacionCiei;

}
