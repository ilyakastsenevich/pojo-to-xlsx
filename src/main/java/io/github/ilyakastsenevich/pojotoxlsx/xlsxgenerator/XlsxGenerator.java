package io.github.ilyakastsenevich.pojotoxlsx.xlsxgenerator;

import java.util.List;
import java.util.Map;

/**
 * XlsxGenerator is an interface for generating Excel xlsx documents from POJOs.
 * It provides two methods for creating xlsx from a single list of objects or a map of sheet names to lists of objects.
 * It also provides a static factory method for obtaining an instance of the interface.
 */
public interface XlsxGenerator {

    /**
     * Generates an Excel xlsx document from a list of objects.
     * Each object represents a row.
     * Object fields' values are read through reflection.
     * The first row of the xlsx will contain the names of the object's fields as column headers.
     * The subsequent rows will contain the values of the fields for each object in the list.
     *
     * @param dataList the list of objects to be written to the xlsx
     *
     * @return a byte array representing the xlsx document or empty byte array if dataList is empty
     */
    byte[] generateXlsx(List<?> dataList);

    /**
     * Generates an Excel xlsx document from a list of objects with specified Sheet names.
     * Each object in the list represents a row.
     * Object fields' values are read through reflection.
     * The first row of the xlsx will contain the names of the object's fields as column headers.
     * The subsequent rows will contain the values of the fields for each object in the list.
     *
     * @param sheetNameToDataListMap the map of Sheet names and the list of objects to be written to the sheet of xlsx
     *
     * @return a byte array representing the xlsx document or empty byte array if sheetNameToDataListMap is empty
     */
    byte[] generateXlsx(Map<String, List<?>> sheetNameToDataListMap);

    /**
     * Returns an instance of XlsxGenerator implementation.
     * @return a XlsxGenerator object
     */
    static XlsxGenerator getInstance() {
        return new XlsxGeneratorImpl();
    }

}
