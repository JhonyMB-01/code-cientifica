package contancia.config;

import jakarta.enterprise.context.ApplicationScoped;
import lombok.Getter;
import org.eclipse.microprofile.config.inject.ConfigProperty;

@Getter
@ApplicationScoped
public class DocumentConfig {

    @ConfigProperty(name = "document.excel.path")
    String excelPath;

    @ConfigProperty(name = "document.word.template")
    String wordTemplate;

    @ConfigProperty(name = "document.libreoffice.path")
    String libreOfficePath;

    @ConfigProperty(name = "documentos.base-path")
    String basePath;

    @ConfigProperty(name = "documentos.excel")
    String excelFile;

    @ConfigProperty(name = "documentos.plantillas-path")
    String plantillasPath;

    @ConfigProperty(name = "documentos.plantilla-extension")
    String plantillaExtension;


}
