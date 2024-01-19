package io.github.ilyakastsenevich.pojotoxlsx.xlsxgenerator;

import io.github.ilyakastsenevich.pojotoxlsx.exception.PojoToXlsxException;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

class XlsxGeneratorImpl implements XlsxGenerator {

    @Override
    public byte[] generateXlsx(List<?> dataList) {
        if (CollectionUtils.isEmpty(dataList)) {
            return new byte[0];
        }

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            createSheet(workbook, null, dataList);
            return toByteArray(workbook);
        } catch (Exception e) {
            throw new PojoToXlsxException(e);
        }
    }

    @Override
    public byte[] generateXlsx(Map<String, List<?>> sheetNameToDataListMap) {
        if (MapUtils.isEmpty(sheetNameToDataListMap)) {
            return new byte[0];
        }

        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            for (Map.Entry<String, List<?>> entry : sheetNameToDataListMap.entrySet()) {
                createSheet(workbook, entry.getKey(), entry.getValue());
            }
            return toByteArray(workbook);
        } catch (Exception e) {
            throw new PojoToXlsxException(e);
        }
    }

    private void createSheet(XSSFWorkbook workbook, String sheetName, List<?> dataList) {
        XSSFSheet sheet = sheetName != null ? workbook.createSheet(sheetName) : workbook.createSheet();

        if (CollectionUtils.isEmpty(dataList)) {
            return;
        }

        Class<?> dataListElementClass = getClassOfFirstElement(dataList);
        createHeaders(sheet, dataListElementClass);

        XSSFRow row;

        int rowNumber = 1;

        for (Object object : dataList) {
            row = sheet.createRow(rowNumber);

            int cellNumber = 0;

            Cell numberCell = row.createCell(cellNumber++);
            numberCell.setCellValue(rowNumber++);

            for (Field field : object.getClass().getDeclaredFields()) {
                Cell cell = row.createCell(cellNumber++);
                String value = null;

                try {
                    value = String.valueOf(FieldUtils.readDeclaredField(object, field.getName(), true));
                } catch (IllegalAccessException e) {
                    throw new PojoToXlsxException(e);
                }

                cell.setCellValue(value);
            }

        }

        if (sheet.getPhysicalNumberOfRows() > 0) {
            Row sheetRow = sheet.getRow(sheet.getFirstRowNum());
            Iterator<Cell> cellIterator = sheetRow.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                Integer columnIndex = cell.getColumnIndex();
                sheet.autoSizeColumn(columnIndex);
            }
        }
    }

    private void createHeaders(XSSFSheet spreadsheet, Class<?> clazz) {
        XSSFRow row = spreadsheet.createRow(0);

        int cellNumber = 0;
        Cell numberCell = row.createCell(cellNumber++);
        numberCell.setCellValue("№");

        Field[] fields = clazz.getDeclaredFields();

        for (Field field : fields) {
            Cell cell = row.createCell(cellNumber++);
            cell.setCellValue(field.getName());
        }
    }

    private byte[] toByteArray(XSSFWorkbook workbook) {
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            throw new PojoToXlsxException(e);
        }
    }

    private Class<?> getClassOfFirstElement(List<?> list) {
        if (CollectionUtils.isEmpty(list)) {
            return null;
        }
        return list.get(0).getClass();
    }
}
