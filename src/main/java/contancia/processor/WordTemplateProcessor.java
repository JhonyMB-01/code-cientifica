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


    /**
     * Procesa una lista de párrafos buscando placeholders
     * con formato ${NombrePlaceholder}.
     *
     * @param paragraphs lista de párrafos del documento
     * @param datos      mapa con los valores de los placeholders
     */
    public void procesarParrafos(
            List<XWPFParagraph> paragraphs,
            Map<String, String> datos) {

        if (paragraphs == null || paragraphs.isEmpty()) {
            return;
        }

        for (XWPFParagraph paragraph : new ArrayList<>(paragraphs)) {

            if (paragraph == null) {
                continue;
            }

            procesarParrafo(
                    paragraph,
                    datos
            );
        }
    }


    /**
     * Procesa un párrafo individual.
     */
    private void procesarParrafo(
            XWPFParagraph paragraph,
            Map<String, String> datos) {

        if (paragraph == null) {
            return;
        }

        if (datos == null || datos.isEmpty()) {
            return;
        }

        List<XWPFRun> runs = paragraph.getRuns();

        if (runs == null || runs.isEmpty()) {
            return;
        }


        /*
         * Construir el texto completo del párrafo.
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

            /*
             * Si el placeholder no existe en el mapa,
             * no se modifica.
             */
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
         * Procesar desde el final hacia el inicio.
         *
         * Esto evita que los índices de los placeholders
         * posteriores se vean afectados por los reemplazos.
         */
        for (int i = placeholders.size() - 1;
             i >= 0;
             i--) {

            /*
             * El párrafo puede haber sido eliminado por un
             * placeholder anterior.
             */
            if (!parrafoExiste(paragraph)) {
                return;
            }

            reemplazarPlaceholder(
                    paragraph,
                    placeholders.get(i)
            );
        }
    }


    /**
     * Reemplaza un placeholder dentro de un párrafo.
     *
     * Si el valor es vacío y el placeholder representa
     * el contenido completo del párrafo, se elimina
     * completamente el párrafo.
     */
    private void reemplazarPlaceholder(
            XWPFParagraph paragraph,
            Placeholder placeholder) {

        if (paragraph == null || placeholder == null) {
            return;
        }

        List<XWPFRun> runs =
                paragraph.getRuns();

        if (runs == null || runs.isEmpty()) {
            return;
        }


        /*
         * Construir nuevamente el texto actual del párrafo.
         *
         * Esto es importante porque pueden existir varios
         * placeholders dentro del mismo párrafo.
         */
        StringBuilder textoBuilder =
                new StringBuilder();

        for (XWPFRun run : runs) {

            String textoRun =
                    obtenerTextoRun(run);

            if (textoRun != null) {
                textoBuilder.append(textoRun);
            }
        }

        String textoActual =
                textoBuilder.toString();


        /*
         * Si el valor está vacío, verificar si el placeholder
         * representa realmente todo el contenido del párrafo.
         *
         * Ejemplo:
         *
         * ${ParrafoDentroUniversidad}
         *
         * Se elimina el párrafo completo.
         *
         * Pero:
         *
         * Texto ${Placeholder} adicional
         *
         * NO elimina el párrafo.
         */
        if (placeholder.valor == null ||
                placeholder.valor.trim().isEmpty()) {

            String textoSinPlaceholder =
                    textoActual.substring(
                            0,
                            Math.min(
                                    placeholder.inicio,
                                    textoActual.length()
                            )
                    )
                            +
                            textoActual.substring(
                                    Math.min(
                                            placeholder.fin,
                                            textoActual.length()
                                    )
                            );

            if (textoSinPlaceholder.trim().isEmpty()) {

                eliminarParrafo(paragraph);

                return;
            }
        }


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


            /*
             * Encontrar Run donde comienza
             * el placeholder.
             */
            if (runInicio == -1 &&
                    placeholder.inicio >= inicio &&
                    placeholder.inicio < fin) {

                runInicio = i;

                offsetInicio =
                        placeholder.inicio -
                                inicio;
            }


            /*
             * Encontrar Run donde termina
             * el placeholder.
             */
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
         * Vaciar Runs intermedios.
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


    /**
     * Elimina completamente un párrafo del documento Word.
     *
     * Esto es diferente a colocar simplemente "":
     * al eliminar el párrafo evitamos que Word conserve
     * el espacio vertical correspondiente al párrafo vacío.
     */
    private void eliminarParrafo(
            XWPFParagraph paragraph) {

        if (paragraph == null) {
            return;
        }

        try {

            var ctp =
                    paragraph.getCTP();

            if (ctp == null) {
                return;
            }

            var parent =
                    ctp.getDomNode().getParentNode();

            if (parent != null) {

                parent.removeChild(
                        ctp.getDomNode()
                );
            }

        } catch (Exception e) {

            /*
             * Como alternativa, si por alguna razón
             * no se puede eliminar el nodo XML,
             * vaciamos los Runs del párrafo.
             */
            List<XWPFRun> runs =
                    paragraph.getRuns();

            if (runs != null) {

                for (XWPFRun run : runs) {

                    if (run != null) {

                        run.setText(
                                "",
                                0
                        );
                    }
                }
            }
        }
    }


    /**
     * Verifica si el párrafo todavía pertenece
     * al documento XML.
     */
    private boolean parrafoExiste(
            XWPFParagraph paragraph) {

        if (paragraph == null) {
            return false;
        }

        try {

            return paragraph.getCTP() != null &&
                    paragraph.getCTP().getDomNode() != null &&
                    paragraph.getCTP().getDomNode().getParentNode() != null;

        } catch (Exception e) {

            return false;
        }
    }


    /**
     * Obtiene todo el texto contenido dentro
     * de un XWPFRun.
     */
    private String obtenerTextoRun(
            XWPFRun run) {

        if (run == null) {
            return "";
        }

        StringBuilder texto =
                new StringBuilder();

        int cantidad =
                run.getCTR()
                        .sizeOfTArray();

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


    /**
     * Representa un placeholder encontrado
     * dentro del texto del párrafo.
     */
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
