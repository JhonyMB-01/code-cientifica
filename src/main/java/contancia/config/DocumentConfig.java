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


}
