package ua.com.smiddle.excelutil.exception;

/**
 * Created by ksa on 26.03.17.
 */
public class SEUException extends Exception {

    public SEUException(String message) {
        super(message);
    }

    @Override
    public String toString() {
        return getClass().getSimpleName() + ": " + super.getMessage();
    }
}
