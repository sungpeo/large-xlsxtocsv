package com.hulkook.parsing.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

/**
 * Convert from xlsx to csv, using stream
 *
 */
public class FromXLSXToCSV {

    private static final String SAX_PARSER = "org.apache.xerces.parsers.SAXParser";

    public static void main( String[] args ) {
        System.out.println( "Hello World!" );
        new FromXLSXToCSV().convertXlsxToCsv();
        System.out.println("Converting job is done.");
    }

    public void convertXlsxToCsv() {


        try {
            String inFilepath = "E:/Book1.xlsx";
            String outFilepath = "E:/Book1.xlsx.csv";
            OPCPackage pkg = OPCPackage.open(new FileInputStream(new File(inFilepath)));
            XSSFReader r = new XSSFReader(pkg);
            SharedStringsTable sst = r.getSharedStringsTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) r.getSheetsData();
            InputStream sheet = iter.next();

            handleExcelSheet(sst, sheet, new FileOutputStream(new File(outFilepath)), iter.getSheetName());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Handles an individual Excel sheet from the entire Excel document.
     */
    private void handleExcelSheet(SharedStringsTable sst, final InputStream sheetInputStream, final OutputStream csvOutputStream,
                                  String sName) throws IOException {

        try {

            String csvDelimter = ";";
            int startingRowNum = 1;

            XMLReader parser = XMLReaderFactory.createXMLReader(SAX_PARSER);
            ExcelSheetRowHandler handler = new ExcelSheetRowHandler(sst, csvOutputStream, csvDelimter, startingRowNum);
            parser.setContentHandler(handler);
            InputSource sheetSource = new InputSource(sheetInputStream);
            try {
                parser.parse(sheetSource);
                sheetInputStream.close();
            } catch (SAXException se) {
                System.out.println("Error occurred while processing Excel sheet "+handler.getSheetName());
            }


        } catch (SAXException saxE) {
            System.out.println("Failed to create instance of SAXParser "+ SAX_PARSER);
            throw new IOException(saxE);
        } finally {
            sheetInputStream.close();
        }
    }


    /**
     * Extracts every row from an Excel Sheet and generates a corresponding JSONObject whose key is the Excel CellAddress and value
     * is the content of that CellAddress converted to a String
     */
    private class ExcelSheetRowHandler extends DefaultHandler {

        private static final String SAX_CELL_REF = "c";
        private static final String SAX_CELL_TYPE = "t";
        private static final String SAX_CELL_STRING = "s";
        private static final String SAX_CELL_CONTENT_REF = "v";
        private static final String SAX_ROW_REF = "row";
        private static final String SAX_SHEET_NAME_REF = "sheetPr";
        private static final String UNKNOWN_SHEET_NAME = "UNKNOWN";
        private final String OUTPUT_CHARACTER_SET = "UTF-8";

        private SharedStringsTable sst;
        private String currentContent;
        private boolean nextIsString;
        final private OutputStream outputStream;
        private boolean firstColInRow;
//        long rowCount;
        String sheetName;
        final String delimiterForCell;
        final long startingRowNum;
        long currentRow;
        boolean didWriteContent;
        private CellAddress cellRef = null;
        private int distance = 0;
        private int columnCountInFirstRow = 0;
        private int currentColumn= 0;

        private ExcelSheetRowHandler(SharedStringsTable sst, OutputStream outputStream, String delimiterForCell, long startingRowNum) {
            this.sst = sst;
            this.outputStream = outputStream;
            this.firstColInRow = true;
//            this.rowCount = 0l;
            this.sheetName = UNKNOWN_SHEET_NAME;
            this.delimiterForCell = delimiterForCell;
            this.startingRowNum = startingRowNum;

        }
        Logger logger = LoggerFactory.getLogger(getClass().getName());

        private Logger getLogger() {
            return this.logger;
        }


        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            if (name.equals(SAX_CELL_REF)) {
                didWriteContent = false;

            if(firstColInRow) {
                addMissedDelimitersIfItIsNotACell(attributes);
            }
            countDistanceOfCells(attributes);
            String cellType = attributes.getValue(SAX_CELL_TYPE);

            //This parser can convert SSTINDEX cell type only. (except for BOOL, DATE, DATETIME, TIME, FORMULA, NUMBER)
            if(cellType != null && cellType.equals(SAX_CELL_STRING)) {
                nextIsString = true;

            } else {
                    nextIsString = false;
            }

        } else if (name.equals(SAX_ROW_REF)) {

            firstColInRow = true;
            this.currentRow = Long.valueOf(attributes.getValue(0));

        } else if (name.equals(SAX_SHEET_NAME_REF)) {
            //it can be null
            sheetName = attributes.getValue(0) != null ? attributes.getValue(0) : UNKNOWN_SHEET_NAME;
        }
        currentContent = "";
    }

    @Override
    public void endElement(String uri, String localName, String name) throws SAXException {

        if(currentRow < startingRowNum) {
            if (firstColInRow) {
                firstColInRow = false;
            }
            return;

        } if (nextIsString) {
            int idx = Integer.parseInt(currentContent);
            currentContent = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
            nextIsString = false;

        } if(name.equals(SAX_CELL_REF) && !didWriteContent) {
            try {
                countColumns();
                outputStream.write(this.delimiterForCell.getBytes(OUTPUT_CHARACTER_SET));

            } catch (IOException e) {
                getLogger().error("IO error encountered while writing content of parsed cell " + "value from sheet {}", new Object[]{getSheetName()}, e);
                throw new SAXException("noContent", e);

            }
        }
        if (name.equals(SAX_CELL_CONTENT_REF)) {
            didWriteContent = true;

            if (firstColInRow) {
                firstColInRow = false;

                try {
                    countColumns();
                    outputStream.write(currentContent.getBytes(OUTPUT_CHARACTER_SET));

                } catch (IOException e) {
                    getLogger().error("IO error encountered while writing content of parsed cell " + "value from sheet {}", new Object[]{getSheetName()}, e);
                    throw new SAXException("firstColInRow:true", e);

                }
            } else {
                try {
                    addMissedDelimtersForBlankCell(outputStream);
                    countColumns();
                    outputStream.write((this.delimiterForCell + currentContent).getBytes(OUTPUT_CHARACTER_SET));


                } catch (IOException e) {
                    getLogger().error("IO error encountered while writing content of parsed cell " + "value from sheet {}", new Object[]{getSheetName()}, e);
                    throw new SAXException("firstColInRow:false", e);
                }
            }
        }
        if (name.equals(SAX_ROW_REF)) {
            //If this is the first row and the end of the row element has been encountered then that means no columns were present.
            if (!firstColInRow) {
                try {
                    addMissedDelimitersAtTheEnd(outputStream);
//                    rowCount++;
                    outputStream.write("\n".getBytes(OUTPUT_CHARACTER_SET));
                } catch (IOException e) {
                    getLogger().error("IO error encountered while writing new line indicator", e);
                }
            }
        }
    }
        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            currentContent += new String(ch, start, length);
        }

        public String getSheetName() {
            return sheetName;
        }

        private void countColumns(){
            currentColumn++;

            if(this.currentRow==1){
                columnCountInFirstRow++;
            }
        }

        private void addMissedDelimitersIfItIsNotACell (Attributes attributes) {
            CellAddress currentCellRef = new CellAddress(attributes.getValue("r"));
            CellAddress aCell = new CellAddress(currentCellRef.getRow(), CellReference.convertColStringToIndex("A"));
            int distanceFromA = currentCellRef.compareTo(aCell);
            while(distanceFromA-->0){
                try {
                    countColumns();
                    outputStream.write(this.delimiterForCell.getBytes(OUTPUT_CHARACTER_SET));
                } catch (IOException e) {
                    getLogger().warn("addMissedDelimitersIfItIsNotACell", e);
                }
            }
        }

        private void countDistanceOfCells(Attributes attributes){
            CellAddress currentCellRef = new CellAddress(attributes.getValue("r"));
            if(cellRef!=null) {
                this.distance = currentCellRef.compareTo(cellRef);
            } else {
                this.distance = 0;
            }

            cellRef = currentCellRef;
        }

        private void addMissedDelimtersForBlankCell(OutputStream outputStream) throws IOException {
            while(this.distance-->1){
                countColumns();
                outputStream.write((this.delimiterForCell).getBytes(OUTPUT_CHARACTER_SET));
            }
        }

        private void addMissedDelimitersAtTheEnd(OutputStream outputStream) throws IOException {
            int restOfColumns = columnCountInFirstRow - currentColumn;
            while(restOfColumns-->0) {
                outputStream.write((this.delimiterForCell).getBytes(OUTPUT_CHARACTER_SET));
            }
            currentColumn=0;
        }
    }
}
