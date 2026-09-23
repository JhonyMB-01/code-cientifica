package contancia.processor;

import jakarta.enterprise.context.ApplicationScoped;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@ApplicationScoped
public class WordTemplateProcessor {

    private static final Pattern PLACEHOLDER_PATTERN =
            Pattern.compile("\\$\\{([^}]+)}");


    public void procesarParrafos(
            List<XWPFParagraph> paragraphs,
            Map<String, String> datos) {

        if (paragraphs == null ||
                paragraphs.isEmpty()) {

            return;
        }

        for (XWPFParagraph paragraph :
                paragraphs) {

            procesarParrafo(
                    paragraph,
                    datos
            );
        }
    }


    private void procesarParrafo(
            XWPFParagraph paragraph,
            Map<String, String> datos) {

        if (paragraph == null) {
            return;
        }

        List<XWPFRun> runs =
                paragraph.getRuns();

        if (runs == null ||
                runs.isEmpty()) {

            return;
        }


        /*
         * Construir texto completo.
         */
        StringBuilder textoCompleto =
                new StringBuilder();

        for (XWPFRun run : runs) {

            String texto =
                    obtenerTextoRun(run);

            if (texto != null) {
                textoCompleto.append(texto);
            }
        }

        String texto =
                textoCompleto.toString();

        if (texto.isEmpty()) {
            return;
        }


        /*
         * Buscar placeholders.
         */
        Matcher matcher =
                PLACEHOLDER_PATTERN.matcher(texto);

        List<Placeholder> placeholders =
                new ArrayList<>();

        while (matcher.find()) {

            String clave =
                    matcher.group(1);

            if (!datos.containsKey(clave)) {
                continue;
            }

            String valor =
                    datos.get(clave);

            if (valor == null) {
                valor = "";
            }

            placeholders.add(
                    new Placeholder(
                            matcher.start(),
                            matcher.end(),
                            valor
                    )
            );
        }


        if (placeholders.isEmpty()) {
            return;
        }


        /*
         * Procesar desde el final.
         */
        for (int i =
             placeholders.size() - 1;
             i >= 0;
             i--) {

            reemplazarPlaceholder(
                    paragraph,
                    placeholders.get(i)
            );
        }
    }


    private void reemplazarPlaceholder(
            XWPFParagraph paragraph,
            Placeholder placeholder) {

        List<XWPFRun> runs =
                paragraph.getRuns();

        int posicion = 0;

        int runInicio = -1;
        int runFin = -1;

        int offsetInicio = -1;
        int offsetFin = -1;


        /*
         * Encontrar Run inicial y final.
         */
        for (int i = 0;
             i < runs.size();
             i++) {

            String textoRun =
                    obtenerTextoRun(
                            runs.get(i)
                    );

            if (textoRun == null) {
                textoRun = "";
            }

            int inicio = posicion;
            int fin =
                    posicion +
                            textoRun.length();


            if (runInicio == -1 &&
                    placeholder.inicio >= inicio &&
                    placeholder.inicio < fin) {

                runInicio = i;

                offsetInicio =
                        placeholder.inicio -
                                inicio;
            }


            if (placeholder.fin > inicio &&
                    placeholder.fin <= fin) {

                runFin = i;

                offsetFin =
                        placeholder.fin -
                                inicio;

                break;
            }


            posicion = fin;
        }


        if (runInicio == -1 ||
                runFin == -1) {

            return;
        }


        XWPFRun runInicial =
                runs.get(runInicio);

        String textoInicial =
                obtenerTextoRun(runInicial);

        if (textoInicial == null) {
            textoInicial = "";
        }


        /*
         * Placeholder dentro de un solo Run.
         */
        if (runInicio == runFin) {

            String antes =
                    textoInicial.substring(
                            0,
                            offsetInicio
                    );

            String despues =
                    textoInicial.substring(
                            offsetFin
                    );

            runInicial.setText(
                    antes +
                            placeholder.valor +
                            despues,
                    0
            );

            return;
        }


        /*
         * Placeholder dividido entre varios Run.
         */
        String antes =
                textoInicial.substring(
                        0,
                        offsetInicio
                );


        XWPFRun runFinal =
                runs.get(runFin);

        String textoFinal =
                obtenerTextoRun(runFinal);

        if (textoFinal == null) {
            textoFinal = "";
        }


        String despues =
                textoFinal.substring(
                        offsetFin
                );


        /*
         * Conservamos el formato del Run inicial.
         */
        runInicial.setText(
                antes +
                        placeholder.valor,
                0
        );


        /*
         * Conservar texto después del placeholder.
         */
        if (!despues.isEmpty()) {

            runFinal.setText(
                    despues,
                    0
            );

        } else {

            runFinal.setText(
                    "",
                    0
            );
        }


        /*
         * Vaciar Run intermedios.
         */
        for (int i =
             runInicio + 1;
             i < runFin;
             i++) {

            runs.get(i).setText(
                    "",
                    0
            );
        }
    }


    private String obtenerTextoRun(
            XWPFRun run) {

        if (run == null) {
            return "";
        }

        StringBuilder texto =
                new StringBuilder();

        int cantidad =
                run.getCTR().sizeOfTArray();

        for (int i = 0;
             i < cantidad;
             i++) {

            String valor =
                    run.getText(i);

            if (valor != null) {
                texto.append(valor);
            }
        }

        return texto.toString();
    }


    private static class Placeholder {

        private final int inicio;
        private final int fin;
        private final String valor;


        private Placeholder(
                int inicio,
                int fin,
                String valor) {

            this.inicio = inicio;
            this.fin = fin;
            this.valor = valor;
        }
    }
}