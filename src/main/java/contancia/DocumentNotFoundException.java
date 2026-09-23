package contancia;

public class DocumentNotFoundException
        extends RuntimeException {

    public DocumentNotFoundException(
            String message) {

        super(message);
    }
}
