package org.example;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Set;

public class Main {
    public static void main(String[] args) throws IOException {
        HSSFWorkbook wb_sozv = readWorkbook("obsh.xls");
        HSSFWorkbook wb_signal = readWorkbook("dmv.xls");
        ArrayList <Element> spisok_sozv  = readPKI_sozv(wb_sozv);
        ArrayList <Element> spisok_signal = readPKI_signal(wb_signal);
        //System.out.println(spisok_sozv);
        //System.out.println("\n\n");
        //System.out.println(spisok_signal);
        ArrayList <Result> result_list = makeResult(spisok_sozv,spisok_signal);
        writeIntoExcel("result.xls",result_list,spisok_signal);
        System.out.println(spisok_signal);
    }

    public static void writeIntoExcel(String file,ArrayList<Result> result_list, ArrayList<Element> signal) throws FileNotFoundException, IOException {
        Workbook book = new HSSFWorkbook();
        Sheet sheet = book.createSheet("DMV2");
        int sdvig = 0;
        for(int i=0; i < result_list.size(); i++){
            Row row = sheet.createRow(sdvig);
            Cell name1 = row.createCell(0);
            name1.setCellValue(result_list.get(i).getSozv().getName());
            Cell count1 = row.createCell(1);
            Cell coeff = row.createCell(4);
            coeff.setCellValue(result_list.get(i).getCoffient());
            count1.setCellValue(result_list.get(i).getSozv().getCount());
            for(int j=0; j < result_list.get(i).getSignal().size(); j++){
                Cell name2 = row.createCell(2);
                name2.setCellValue(result_list.get(i).getSignal().get(j).getName());
                Cell count2 = row.createCell(3);
                count2.setCellValue(result_list.get(i).getSignal().get(j).getCount());
                sdvig++;
                row = sheet.createRow(sdvig);
            }
        }

        for(int i=0; i < signal.size(); i++){
            Row row = sheet.createRow(i+3+sdvig);
            Cell name2 = row.createCell(2);
            name2.setCellValue(signal.get(i).getName());
            Cell count2 = row.createCell(3);
            count2.setCellValue(signal.get(i).getCount());
        }
//        // Нумерация начинается с нуля
//        Row row = sheet.createRow(0);
//
//        // Мы запишем имя и дату в два столбца
//        // имя будет String, а дата рождения --- Date,
//        // формата dd.mm.yyyy
//        Cell name = row.createCell(0);
//        name.setCellValue("John");
//
//        Cell birthdate = row.createCell(1);
//
//        DataFormat format = book.createDataFormat();
//        CellStyle dateStyle = book.createCellStyle();
//        dateStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
//        birthdate.setCellStyle(dateStyle);
//
//
//        // Нумерация лет начинается с 1900-го
//        birthdate.setCellValue(new Date(110, 10, 10));

        // Меняем размер столбца
        for (int i = 0; i < 4; i++){
            sheet.autoSizeColumn(i);
        }
        //sheet.autoSizeColumn(1);

        // Записываем всё в файл
        book.write(new FileOutputStream(file));
        book.close();
    }

    public static ArrayList <Result> makeResult (ArrayList<Element> sozv, ArrayList<Element> signal){
        ArrayList<Result> result_list = new ArrayList<>();
        for (Element e_sozv:sozv){
            String name_sozv = e_sozv
                    .getName()
                    .toLowerCase()
                    .replaceAll("\\s", "")
                    .replaceAll("\\.", "")
                    .replaceAll("-", "")
                    .replaceAll(",", "");
            String name_signal = signal.get(0)
                    .getName()
                    .toLowerCase()
                    .replaceAll("\\s", "")
                    .replaceAll("\\.", "")
                    .replaceAll("-", "")
                    .replaceAll(",", "");
            int min_distance = StringUtils.getLevenshteinDistance(name_sozv, name_signal);
            Element element_exit = signal.get(0);
            for (Element e_signal:signal){
                name_signal = e_signal
                        .getName()
                        .toLowerCase()
                        .replaceAll("\\s", "")
                        .replaceAll("\\.", "")
                        .replaceAll("-", "")
                        .replaceAll(",", "");
                int distance = StringUtils.getLevenshteinDistance(name_sozv, name_signal);
                System.out.println(name_signal);
                if (distance < min_distance){
                    min_distance = distance;
                    element_exit = e_signal;
                }
            }

            ArrayList <Element> dubList = new ArrayList<>();
            for (Element e_signal:signal) {
                name_signal = e_signal
                        .getName()
                        .toLowerCase()
                        .replaceAll("\\s", "")
                        .replaceAll("\\.", "")
                        .replaceAll("-", "")
                        .replaceAll(",", "");
                int distance = StringUtils.getLevenshteinDistance(name_sozv, name_signal);
                System.out.println(name_signal);
                if (distance == min_distance) {
                    dubList.add(e_signal);
                }
            }

            //signal.remove(element_exit);
            System.out.println(min_distance);
            Result result_para = new Result(e_sozv,dubList,min_distance);
            result_list.add(result_para);
        }
        for (Result result : result_list){
            for (Element element : result.getSignal()) {
                signal.remove(element);
            }
        }
        return result_list;
    }

    public static ArrayList <Element> readPKI_sozv(HSSFWorkbook wb){
        ArrayList<Element> spisok = new ArrayList<>();
        if (wb != null) {
            HSSFSheet sheet = wb.getSheet("ПРМД ДМВ2 464512.071");
            Iterator<Row> rowIter = sheet.rowIterator();
            rowIter.next();
            while (rowIter.hasNext()) {
                StringBuilder sb = new StringBuilder();
                HSSFRow row = (HSSFRow) rowIter.next();
                HSSFCell type = row.getCell(0);
                String str0 = type.getRichStringCellValue().getString();
                System.out.println(str0);
                HSSFCell name = row.getCell(1);
                String str1 = name.getRichStringCellValue().getString();
                System.out.println(str1);
                HSSFCell docum = row.getCell(2);
                String str2 = docum.getRichStringCellValue().getString();
                System.out.println(str2);
                HSSFCell count = row.getCell(3);
                //int str3 = (int)Double.parseDouble(count.getRichStringCellValue().getString());
                int str3 = (int)Double.parseDouble(String.valueOf(count.getNumericCellValue()));
                System.out.println(count+"\n");
                sb.append(str0);
                sb.append(" ");
                sb.append(str1);
                sb.append(" ");
                sb.append(str2);
                Element element = new Element(sb.toString(),str3);
                spisok.add(element);
            }

        }
        return spisok;
    }

    public static ArrayList <Element> readPKI_signal(HSSFWorkbook wb){
        ArrayList<Element> spisok = new ArrayList<>();
        //System.out.println("111");
        if (wb != null) {
            HSSFSheet sheet = wb.getSheet("Лист1");
            Iterator<Row> rowIter = sheet.rowIterator();
            rowIter.next();
            rowIter.next();
            while (rowIter.hasNext()) {
                StringBuilder sb = new StringBuilder();
                HSSFRow row = (HSSFRow) rowIter.next();
                HSSFCell name = row.getCell(1);
                String str1 = name.getRichStringCellValue().getString();
                HSSFCell count = row.getCell(2);
                int str3 = (int)Double.parseDouble(String.valueOf(count.getNumericCellValue()));
                sb.append(str1);
                Element element = new Element(sb.toString(),str3);
                spisok.add(element);
            }

        }
        return spisok;
    }


    /**
     * Метод возвращает объект класса XSSFWorkbook
     * если все удачно и null в другом случае.
     * @param filename имя таблици
     * @return таблица
     */
    public static HSSFWorkbook readWorkbook(String filename) {
        try {
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(filename));
            return wb;
        }
        catch (Exception e) {
            return null;
        }
    }
}