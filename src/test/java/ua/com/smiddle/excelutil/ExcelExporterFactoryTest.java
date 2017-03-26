package ua.com.smiddle.excelutil;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Random;

/**
 * Created by ksa on 26.03.17.
 */
public class ExcelExporterFactoryTest {
    private final SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
    private String path = "//home//ksa//test//";
    private Configurer conf;
    private List<Object[]> tableHeader;
    private List<Object[]> data;


    @org.junit.Before
    public void setUp() throws Exception {
        conf = Configurer.buildNewConfigurer()
                .showReportName(false);
        fillTableHeader();
        fillData();

    }

    @org.junit.After
    public void tearDown() throws Exception {
    }

    @org.junit.Test
    public void buildNewInstance() throws Exception {
        ExcelExporter exporter = ExcelExporterFactory.buildNewInstance(conf);
        exporter.buildDocument(tableHeader, data);
        FileOutputStream os = new FileOutputStream(new File(path + this.getClass().getSimpleName() +
                "-" + format.format(new Date()) + ".xls"));
        exporter.writeDocument(os);
        os.close();
    }

    private void fillTableHeader() {
        tableHeader = new ArrayList<>();
        Object[] row = new Object[4];
        row[0] = "Integer";
        row[1] = "Double";
        row[2] = "Data";
        row[3] = "String";
        tableHeader.add(row);
    }

    private void fillData() {
        data = new ArrayList<>();
        Random r = new Random();
        for (int i = 0; i < 10; i++) {
            Object[] row = new Object[4];
            row[0] = r.nextInt(100000);
            row[1] = r.nextDouble();
            row[2] = new Date(System.currentTimeMillis());
            row[3] = String.valueOf("String: " + System.currentTimeMillis());
            data.add(row);
        }
    }

}