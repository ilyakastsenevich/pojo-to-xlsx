package io.github.ilyakastsenevich.pojotoxlsx.xlsxgenerator;


import lombok.AllArgsConstructor;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

class XlsxGeneratorImplTest {

    @Test
    void generateXlsx() throws IOException {
        //create your input object, it can be of any class
        YourPojo row1 = new YourPojo("name1", "value1",  1);
        YourPojo row2 = new YourPojo("name2", "value2",  2);
        YourPojo row3 = new YourPojo("name3", "value3",  3);

        XlsxGenerator xlsxGenerator = XlsxGenerator.getInstance();

        byte[] result = xlsxGenerator.generateXlsx(List.of(row1, row2, row3));

        File outputFile = new File("src/test/resources/result.xlsx");
        Files.write(outputFile.toPath(), result);
    }

    @Test
    void generateXlsxWithSheetname() throws IOException {
        //create your input object, it can be of any class
        YourPojo row1 = new YourPojo("name1", "value1",  1);
        YourPojo row2 = new YourPojo("name2", "value2",  2);
        YourPojo row3 = new YourPojo("name3", "value3",  3);

        //result xlsx will contain 2 sheets
        String sheetName1 = "MySheetName1";
        String sheetName2 = "MySheetName2";

        XlsxGenerator xlsxGenerator = XlsxGenerator.getInstance();

        Map<String, List<?>> inputMap = new LinkedHashMap<>();
        inputMap.put(sheetName1, List.of(row1, row2, row3));
        inputMap.put(sheetName2, List.of(row1, row2, row3));

        byte[] result = xlsxGenerator.generateXlsx(inputMap);

        File outputFile = new File("src/test/resources/result2.xlsx");
        Files.write(outputFile.toPath(), result);
    }

    @AllArgsConstructor
    private static class YourPojo {
        // fields' names will become columns' headers
        // values will become cells' values
        private String name;
        private String valueText;
        private Integer valueInt;
    }
}