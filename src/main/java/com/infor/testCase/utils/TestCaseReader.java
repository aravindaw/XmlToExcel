package com.infor.testCase.utils;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.logging.Logger;

public class TestCaseReader {

    private static Workbook workbook;
    private static int rowNum;


    private final static int TEST_SUITE_NAME_COLUMN = 0;
    private final static int TEST_CASE_NAME_COLUMN = 1;
    private static final Logger LOGGER = Logger.getLogger(TestCaseReader.class.getName());


    public static void main(String[] args) {
        Scanner scan = new Scanner(System.in);
        System.out.println("Please give XML file path and press ENTER key (E.g. C:\\testscripts\\src\\plan)");
        String inputPath = scan.nextLine();
        inputPath = inputPath.trim();

        prepareExcelSheet();

        //here it can only identify whether it's directory or not. If the are any files other tha Directory or xml there will be exception
        File f = new File(inputPath);
        File[] files = f.listFiles();
        assert files != null;
        for (File file : files) {
            if (file.isDirectory()) {
                try {
                    LOGGER.info(file.getCanonicalPath() + " is a Directory and it's not executable..");
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                try {
                    String path = file.getAbsolutePath();
                    LOGGER.info(String.valueOf(file));
                    String fileName = file.getName();
                    writeExcel(fileName, path);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static void writeExcel(String fileName, String path) throws Exception {
        File xmlFile = new File(path);
        Sheet sheet = workbook.getSheetAt(0);
        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);

        NodeList nList = doc.getElementsByTagName("class");
        for (int i = 0; i < nList.getLength(); i++) {
            LOGGER.info("Processing element " + (i + 1) + "/" + nList.getLength());
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                String testCaseName = element.getAttribute("name");

                Row row = sheet.createRow(rowNum++);
                Cell cell = row.createCell(TEST_SUITE_NAME_COLUMN);
                cell.setCellValue(fileName);

                cell = row.createCell(TEST_CASE_NAME_COLUMN);
                cell.setCellValue(testCaseName.replaceAll("scripts.", ""));
            }
        }

        FileOutputStream fileOut = new FileOutputStream("TestCases.xlsx");
        workbook.write(fileOut);
        LOGGER.info("Write Excel is finished, processed " + nList.getLength() + " substances in" + fileName);
        LOGGER.info("==========================================================");

            //##If you wanna delete the excel sheet after read use this block##
//        if (xmlFile.exists()) {
//            System.out.println("delete file-> " + xmlFile.getAbsolutePath());
//            if (!xmlFile.delete()) {
//                System.out.println("file '" + xmlFile.getAbsolutePath() + "' was not deleted!");
//            }
//        }
    }

    private static void prepareExcelSheet() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(TEST_SUITE_NAME_COLUMN);
        cell.setCellValue("Test Suite name");
        cell.setCellStyle(style);

        cell = row.createCell(TEST_CASE_NAME_COLUMN);
        cell.setCellValue("Test Case name");
        cell.setCellStyle(style);
    }
}