package ua.com.smiddle.excelutil;

/**
 * @author ksa on 27.03.17.
 * @project SmiddleExcelUtil
 */
public class tst {
    public static void main(String[] args) {
        Integer a = 10;
        Double b = Double.valueOf(a);
        Number n = new Byte((byte) 1);

        if(b instanceof Number)
            System.out.println("yes");

    }
}
