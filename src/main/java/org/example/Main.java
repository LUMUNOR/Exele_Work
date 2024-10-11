package org.example;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.Set;

public class Main {
    public static void main(String[] args) {
        XSSFWorkbook wb = readWorkbook("Capacitors.xlsx");
        Set<String> manufacturers = readManufacturers(wb);
    }

    /**
     * Создание файлов по производителям
     * @param wb
     * @param manufacturers
     */
    public static void makeManfBook (XSSFWorkbook wb, Set<String> manufacturers) {

    }
    /**
     * Метод достёт список производителей
     * @param wb
     * @return
     */
    public static Set<String> readManufacturers(XSSFWorkbook wb){
        Set<String> manufacturers = new LinkedHashSet<>();
        if (wb != null) {
            XSSFSheet sheet = wb.getSheet("Лист1");
            Iterator rowIter = sheet.rowIterator();
            rowIter.next();
            while (rowIter.hasNext()) {
                XSSFRow row = (XSSFRow) rowIter.next();
                XSSFCell manufacturer = row.getCell(4);
                String str = manufacturer.getRichStringCellValue().getString();
                if (str == null) manufacturers.add("Без производителя");
                else manufacturers.add(str);
            }

        }
        return manufacturers;
    }

    /**
     * Метод возвращает объект класса XSSFWorkbook
     * если все удачно и null в другом случае.
     * @param filename имя таблици
     * @return таблица
     */
    public static XSSFWorkbook readWorkbook(String filename) {
        try {
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(filename));
            return wb;
        }
        catch (Exception e) {
            return null;
        }
    }
}