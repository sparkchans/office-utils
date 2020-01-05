package word;

/**
 * @author sparkchan
 * @date 2020/1/5
 */
public class NotWordTemplateReplaceableException extends RuntimeException {

    private static final long serialVersionUID = -5928973167286410875L;

    public NotWordTemplateReplaceableException() {
        super();
    }

    public NotWordTemplateReplaceableException(String message) {
        super(message);
    }
}
