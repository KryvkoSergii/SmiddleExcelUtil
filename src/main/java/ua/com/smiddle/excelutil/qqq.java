package ua.com.smiddle.excelutil;

import java.util.Date;

/**
 * Created by ksa on 25.03.17.
 */
public class qqq {
    public static void main(String[] args) {
        Object[] array = new Object[7];
        array[0] = 1L;
        array[1] = 5.5;
        array[2] = "string";
        array[3] = new Date();
        array[4] = true;
        array[5] = 5.5f;
        array[6] = 'a';

        for (Object o : array) {
            Class c = null;
            if (o != null) c = o.getClass();
            System.out.println(c != null ? c.getTypeName() : c);
        }

        System.out.println(Double.MAX_VALUE);
        System.out.println(Long.MAX_VALUE);
        double d = (Double.MAX_VALUE - Long.MAX_VALUE);
        if (Double.MAX_VALUE > Long.MAX_VALUE)
            System.out.println(" double "+d);
        else System.out.println("long");

    }
}
