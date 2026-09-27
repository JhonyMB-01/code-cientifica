package utils;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Locale;

public final class Utils {

    public static String formarterFecha(String fecha) {
        DateTimeFormatter salida = DateTimeFormatter.ofPattern(
                "dd 'de' MMMM 'del' yyyy", Locale.forLanguageTag("es-ES"));

        for (String formato : new String[]{"M/d/yy", "M/d/yyyy"}) {
            try {
                return LocalDate.parse(fecha, DateTimeFormatter.ofPattern(formato))
                        .format(salida);
            } catch (DateTimeParseException ignored) {
            }
        }

        throw new IllegalArgumentException("Fecha inválida: " + fecha);
    }
}
