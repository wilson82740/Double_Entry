package importAction;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelReader extends DefaultHandler {
  int space;

  int maxAscII = 65;
  int thisAscII = 65;
  int pastAscII = 64; 
  Boolean firstposition = true;
  
  private SharedStringsTable sst;

  
  private String lastContents;
  private boolean nextIsString;

  
  private boolean isTElement;
  private boolean isDate;

  List<String> excelCell = new ArrayList<String>();
  List<List<String>> excelRow = new ArrayList<List<String>>();
  List<List<List<String>>> excelSheet = new ArrayList<List<List<String>>>();
  Boolean firstorlast = false;

  /**
   *
   * @param fileStream
   * @throws Exception
   * @return List
   */
  public List<List<List<String>>> processAll(InputStream fileStream) throws Exception {
    OPCPackage pkg = OPCPackage.open(fileStream);
    XSSFReader r = new XSSFReader(pkg);
    SharedStringsTable sst = r.getSharedStringsTable();
    XMLReader parser = fetchSheetParser(sst);
    Iterator<InputStream> sheets = r.getSheetsData();
    int sheetnum = 1;
    while (sheets.hasNext()) {
      excelRow = new ArrayList<List<String>>();
      InputStream sheet = sheets.next();
      if (sheetnum == 1 || !sheets.hasNext()) {
        firstorlast = true;
      }
      InputSource sheetSource = new InputSource(sheet);
      parser.parse(sheetSource);
      sheet.close();
      excelSheet.add(excelRow);
      maxAscII = 65;
      thisAscII = 65;
      pastAscII = 64;
      firstposition = true;
      firstorlast = false;
      sheetnum++;
    }
    return excelSheet;
  }

  /**
   * 讀取XML
   *
   * @param sst
   * @return XMLReader
   * @throws SAXException
   */
  public XMLReader fetchSheetParser(SharedStringsTable sst)
      throws SAXException {
    XMLReader parser =
        XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
    this.sst = sst;
    parser.setContentHandler(this);
    return parser;
  }

  /**
   * 處理XML內容，遇到啟始標籤時處理內容
   *
   * @param uri
   * @param localName
   * @param name
   * @param attributes
   * @throws SAXException
   */
  public void startElement(String uri, String localName, String name,
      Attributes attributes) throws SAXException {

    if ("t".equals(name)) {

    } else if ("c".equals(name)) {
      isDate = false;
      isTElement = true;
      positionSet(attributes.getValue("r"));
      String cellType = attributes.getValue("t");
      String cellstyle = attributes.getValue("s");
      if ("s".equals(cellType)) {
        nextIsString = true;
      } else {
        nextIsString = false;
      }

    }
    lastContents = "";
  }

  /**
   *
   * @param uri
   * @param localName
   * @param name
   * @throws SAXException
   */
  public void endElement(String uri, String localName, String name)
      throws SAXException {

    if (nextIsString) {
      try {
        int idx = Integer.parseInt(lastContents);
        lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
      } catch (Exception e) {

      }
    }
    if (isTElement) {
      String value = lastContents.replaceAll("\u00a0", "");
      if (isTElement) {
        isTElement = false;
      } else if ("v".equals(name)) {
        value = value.equals("") ? " " : value;
      }
      addSpace(thisAscII, pastAscII, true);
      excelCell.add(value);
    }
    if (name.equals("row")) {
      addSpace(maxAscII, thisAscII, false);
      if (!firstorlast) {
        if (excelCell.size() < 6) {
          for (int space = 0; space <= 6 - excelCell.size(); space++) {
            excelCell.add("");
          }
        }
      }
      excelRow.add(excelCell);
      pastAscII = 64;
      firstposition = true;
      excelCell = new ArrayList<String>();
    }

  }

  public void characters(char[] ch, int start, int length) throws SAXException {
    lastContents += new String(ch, start, length);
  }

  /**
   * @param cellPosition
   */
  public void positionSet(String cellPosition) {
    int cellASCII;
    int StrASCII = 0;
    for (int i = 0; i < cellPosition.length(); i++) {
      cellASCII = (int) cellPosition.charAt(i);
      if (cellASCII >= 65) {
        StrASCII += cellASCII;
      }
    }
    if (!firstposition) {
      pastAscII = thisAscII;
    } else {
      firstposition = false;
    }
    thisAscII = StrASCII;

    if (StrASCII > maxAscII) {
      maxAscII = StrASCII;
    }
  }

  /**
   *
   * @param a
   * @param b
   * @param c
   */
  public void addSpace(int a, int b, Boolean c) {
    int addspace = 0;
    if (c) {
      if (a - b > 1) {
        addspace = a - b - 1;
      }
    } else {
      if (a - b > 0) {
        addspace = a - b;
      }
    }

    for (int i = 0; i < addspace; i++) {
      excelCell.add("");
    }
  }
}
