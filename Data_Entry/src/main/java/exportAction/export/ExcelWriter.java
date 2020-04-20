package exportAction.export;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import exporModel.ExportCell;
import exporModel.MergeCell;


public abstract class ExcelWriter {

  private SpreadsheetWriter sw;
  public SharedStringsTable sst;
  
  private Map<String, List<List<ExportCell>>> exportRowsMap;
  private Map<String, List<MergeCell>> mergeCellsMap;
  private Map<String, List<Integer>> columnsWidthMap;
  
  public Map<String, List<List<ExportCell>>> getExportRowsMap() {
    return exportRowsMap;
  }

  public void
      setExportRowsMap(Map<String, List<List<ExportCell>>> exportRowsMap) {
    this.exportRowsMap = exportRowsMap;
  }

  public Map<String, List<MergeCell>> getMergeCellsMap() {
    return mergeCellsMap;
  }

  public void setMergeCellsMap(Map<String, List<MergeCell>> mergeCellsMap) {
    this.mergeCellsMap = mergeCellsMap;
  }

  public Map<String, List<Integer>> getColumnsWidthMap() {
    return columnsWidthMap;
  }

  public void setColumnsWidthMap(Map<String, List<Integer>> columnsWidthMap) {
    this.columnsWidthMap = columnsWidthMap;
  }

  /**
   * 
   * 
   * @param fileName
   * @throws Exception
   */
  public void process(String fileName) throws Exception {

    XSSFWorkbook wb = new XSSFWorkbook();
    Map<String, XSSFCellStyle> styles = createStyles(wb);
    List<File> fileArray = new ArrayList<File>();
    List<String> sheetNameArray = new ArrayList<String>();
    
    sst = wb.getSharedStringSource();

    for(Map.Entry<String, List<List<ExportCell>>> entry : exportRowsMap.entrySet()) {
      XSSFSheet sheet = null;
      switch(entry.getKey()) { 
      case "Patient":
        sheet = wb.createSheet("病人基本資料比較");
        break;
      case "Visit_Date":
        sheet = wb.createSheet("訪視日期比較");
        break;
      case "Form Stauts":
        sheet = wb.createSheet("表單填寫狀態比較");
        break;
      default: 
        sheet = wb.createSheet(entry.getKey());
      }
      String sheetRef = sheet.getPackagePart().getPartName().getName();
      sheetNameArray.add(sheetRef);
      File tmp = File.createTempFile("sheet", ".xml");
      Writer fw = new FileWriter(tmp);
      sw = new SpreadsheetWriter(fw);
      generate(styles, sst, entry.getKey(), entry.getValue());
      fw.close();
      fileArray.add(tmp);
    }
    
    FileOutputStream os = new FileOutputStream("template.xlsx");
    wb.write(os);
    os.close();

    File templateFile = new File("template.xlsx");
    FileOutputStream out = new FileOutputStream(fileName);
    substitute(templateFile, fileArray, sheetNameArray, out);
    out.close();

    System.gc();

    for(File file : fileArray) {
      if(file.isFile() && file.exists()) {
        file.delete();
      }
    }
    
    if (templateFile.isFile() && templateFile.exists()) {
      templateFile.delete();
    }
  }

  private Map<String, XSSFCellStyle> createStyles(XSSFWorkbook wb) {
    Map<String, XSSFCellStyle> styles = new HashMap<String, XSSFCellStyle>();
    XSSFDataFormat fmt = wb.createDataFormat();

XSSFColor color = new XSSFColor();
    
    XSSFCellStyle normal = wb.createCellStyle();
    XSSFFont normalFont = wb.createFont();
    normalFont.setFontName("Calibri");
    normal.setFont(normalFont);
    styles.put("normal", normal);
    
    XSSFCellStyle boldandBorder = wb.createCellStyle();
    XSSFFont boldandBorderFont = wb.createFont();
    boldandBorderFont.setFontName("Calibri");
    boldandBorderFont.setBold(true);
    boldandBorder.setFont(boldandBorderFont);
    boldandBorder.setBorderBottom(BorderStyle.THICK);
    boldandBorder.setBorderTop(BorderStyle.THICK);
    boldandBorder.setBorderLeft(BorderStyle.THICK);
    boldandBorder.setBorderRight(BorderStyle.THICK);
    boldandBorder.setAlignment(HorizontalAlignment.CENTER);
    boldandBorder.setVerticalAlignment(VerticalAlignment.CENTER);
    styles.put("boldandBorder", boldandBorder);
    
    XSSFCellStyle bold = wb.createCellStyle();
    XSSFFont boldFont = wb.createFont();
    boldFont.setFontName("Calibri");
    boldFont.setBold(true);
    bold.setFont(boldFont);
    styles.put("bold", bold);
    
    XSSFCellStyle tlBorder = wb.createCellStyle();
    XSSFFont tlBorderFont = wb.createFont();
    tlBorderFont.setFontName("Calibri");
    tlBorder.setFont(tlBorderFont);
    tlBorder.setBorderTop(BorderStyle.THICK);
    tlBorder.setBorderLeft(BorderStyle.THICK);
    styles.put("tlBorder", tlBorder);
    
    XSSFCellStyle blBorder = wb.createCellStyle();
    XSSFFont blBorderFont = wb.createFont();
    blBorderFont.setFontName("Calibri");
    blBorder.setFont(blBorderFont);
    blBorder.setBorderBottom(BorderStyle.THICK);
    blBorder.setBorderLeft(BorderStyle.THICK);
    styles.put("blBorder", blBorder);
    
    XSSFCellStyle trBorder = wb.createCellStyle();
    XSSFFont trBorderFont = wb.createFont();
    trBorderFont.setFontName("Calibri");
    trBorder.setFont(trBorderFont);
    trBorder.setBorderTop(BorderStyle.THICK);
    trBorder.setBorderRight(BorderStyle.THICK);
    styles.put("trBorder", trBorder);
    
    XSSFCellStyle brBorder = wb.createCellStyle();
    XSSFFont brBorderFont = wb.createFont();
    brBorderFont.setFontName("Calibri");
    brBorder.setFont(brBorderFont);
    brBorder.setBorderBottom(BorderStyle.THICK);
    brBorder.setBorderRight(BorderStyle.THICK);
    styles.put("brBorder", brBorder);
    
    XSSFCellStyle topBorder = wb.createCellStyle();
    XSSFFont topBorderFont = wb.createFont();
    topBorderFont.setFontName("Calibri");
    topBorder.setFont(topBorderFont);
    topBorder.setBorderTop(BorderStyle.THICK);
    styles.put("topBorder", topBorder);
    
    XSSFCellStyle bottomBorder = wb.createCellStyle();
    XSSFFont bottomBorderFont = wb.createFont();
    bottomBorderFont.setFontName("Calibri");
    bottomBorder.setFont(bottomBorderFont);
    bottomBorder.setBorderBottom(BorderStyle.THICK);
    styles.put("bottomBorder", bottomBorder);
    
    XSSFCellStyle yellowtrBorder = wb.createCellStyle();
    XSSFFont yellowtrBorderFont = wb.createFont();
    yellowtrBorderFont.setFontName("Calibri");
    yellowtrBorder.setFont(yellowtrBorderFont);
    yellowtrBorder.setBorderTop(BorderStyle.THICK);
    yellowtrBorder.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("FFFF37");
    yellowtrBorder.setFillForegroundColor(color);
    yellowtrBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("yellowtrBorder", yellowtrBorder);
    
    XSSFCellStyle yellowbrBorder = wb.createCellStyle();
    XSSFFont yellowbrBorderFont = wb.createFont();
    yellowbrBorderFont.setFontName("Calibri");
    yellowbrBorder.setFont(yellowbrBorderFont);
    yellowbrBorder.setBorderBottom(BorderStyle.THICK);
    yellowbrBorder.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("FFFF37");
    yellowbrBorder.setFillForegroundColor(color);
    yellowbrBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("yellowbrBorder", yellowbrBorder);
    
    XSSFCellStyle yellowTopBorder = wb.createCellStyle();
    XSSFFont yellowTopBorderFont = wb.createFont();
    yellowTopBorderFont.setFontName("Calibri");
    yellowTopBorder.setFont(yellowTopBorderFont);
    yellowTopBorder.setBorderTop(BorderStyle.THICK);
    color.setARGBHex("FFFF37");
    yellowTopBorder.setFillForegroundColor(color);
    yellowTopBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("yellowTopBorder", yellowTopBorder);
    
    XSSFCellStyle yellowBottomBorder = wb.createCellStyle();
    XSSFFont yellowBottomBorderFont = wb.createFont();
    yellowBottomBorderFont.setFontName("Calibri");
    yellowBottomBorder.setFont(yellowBottomBorderFont);
    yellowBottomBorder.setBorderBottom(BorderStyle.THICK);
    color.setARGBHex("FFFF37");
    yellowBottomBorder.setFillForegroundColor(color);
    yellowBottomBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("yellowBottomBorder", yellowBottomBorder);
    
    XSSFCellStyle greentrBorder = wb.createCellStyle();
    XSSFFont greentrBorderFont = wb.createFont();
    greentrBorderFont.setFontName("Calibri");
    greentrBorder.setFont(greentrBorderFont);
    greentrBorder.setBorderTop(BorderStyle.THICK);
    greentrBorder.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("93FF93");
    greentrBorder.setFillForegroundColor(color);
    greentrBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("greentrBorder", greentrBorder);
    
    XSSFCellStyle greenbrBorder = wb.createCellStyle();
    XSSFFont greenbrBorderFont = wb.createFont();
    greenbrBorderFont.setFontName("Calibri");
    greenbrBorder.setFont(greenbrBorderFont);
    greenbrBorder.setBorderBottom(BorderStyle.THICK);
    greenbrBorder.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("93FF93");
    greenbrBorder.setFillForegroundColor(color);
    greenbrBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("greenbrBorder", greenbrBorder);
    
    XSSFCellStyle greenTopBorder = wb.createCellStyle();
    XSSFFont greenTopBorderFont = wb.createFont();
    greenTopBorderFont.setFontName("Calibri");
    greenTopBorder.setFont(greenTopBorderFont);
    greenTopBorder.setBorderTop(BorderStyle.THICK);
    color.setARGBHex("93FF93");
    greenTopBorder.setFillForegroundColor(color);
    greenTopBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("greenTopBorder", greenTopBorder);
    
    XSSFCellStyle greenBottomBorder = wb.createCellStyle();
    XSSFFont greenBottomBorderFont = wb.createFont();
    greenBottomBorderFont.setFontName("Calibri");
    greenBottomBorder.setFont(greenBottomBorderFont);
    greenBottomBorder.setBorderBottom(BorderStyle.THICK);
    color.setARGBHex("93FF93");
    greenBottomBorder.setFillForegroundColor(color);
    greenBottomBorder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("greenBottomBorder", greenBottomBorder);
    
    XSSFCellStyle protocolTopic1 = wb.createCellStyle();
    XSSFFont protocolTopic1Font = wb.createFont();
    protocolTopic1Font.setFontName("Calibri");
    protocolTopic1Font.setBold(true);
    protocolTopic1.setFont(protocolTopic1Font);
    protocolTopic1.setBorderBottom(BorderStyle.THICK);
    protocolTopic1.setBorderTop(BorderStyle.THICK);
    protocolTopic1.setBorderLeft(BorderStyle.THICK);
    protocolTopic1.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("93FF93");
    protocolTopic1.setFillForegroundColor(color);
    protocolTopic1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    protocolTopic1.setAlignment(HorizontalAlignment.CENTER);
    protocolTopic1.setVerticalAlignment(VerticalAlignment.CENTER);
    styles.put("protocolTopic1", protocolTopic1);
    
    XSSFCellStyle protocolTopic2 = wb.createCellStyle();
    XSSFFont protocolTopic2Font = wb.createFont();
    protocolTopic2Font.setFontName("Calibri");
    protocolTopic2Font.setBold(true);
    protocolTopic2.setFont(protocolTopic2Font);
    protocolTopic2.setBorderBottom(BorderStyle.THICK);
    protocolTopic2.setBorderTop(BorderStyle.THICK);
    protocolTopic2.setBorderLeft(BorderStyle.THICK);
    protocolTopic2.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("7D7DFF");
    protocolTopic2.setFillForegroundColor(color);
    protocolTopic2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    protocolTopic2.setAlignment(HorizontalAlignment.CENTER);
    protocolTopic2.setVerticalAlignment(VerticalAlignment.CENTER);
    styles.put("protocolTopic2", protocolTopic2);
    
    XSSFCellStyle inconsistentTopic = wb.createCellStyle();
    XSSFFont inconsistentTopicFont = wb.createFont();
    inconsistentTopicFont.setFontName("Calibri");
    inconsistentTopicFont.setBold(true);
    inconsistentTopic.setFont(inconsistentTopicFont);
    inconsistentTopic.setBorderBottom(BorderStyle.THICK);
    inconsistentTopic.setBorderTop(BorderStyle.THICK);
    inconsistentTopic.setBorderLeft(BorderStyle.THICK);
    inconsistentTopic.setBorderRight(BorderStyle.THICK);
    color.setARGBHex("FFD1A4");
    inconsistentTopic.setFillForegroundColor(color);
    inconsistentTopic.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    inconsistentTopic.setAlignment(HorizontalAlignment.CENTER);
    inconsistentTopic.setVerticalAlignment(VerticalAlignment.CENTER);
    styles.put("inconsistentTopic", inconsistentTopic);
    
    XSSFCellStyle valueEmpty = wb.createCellStyle();
    XSSFFont valueEmptyFont = wb.createFont();
    valueEmptyFont.setFontName("Calibri");
    valueEmpty.setFont(valueEmptyFont);
    color.setARGBHex("FFD1A4");
    valueEmpty.setFillForegroundColor(color);
    valueEmpty.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("valueEmpty", valueEmpty);

    XSSFCellStyle header = wb.createCellStyle();
    XSSFFont headerFont = wb.createFont();
    headerFont.setBold(true);
    headerFont.setFontName("Calibri");
    header.setFont(headerFont);
    header.setBorderBottom(BorderStyle.THICK);
    header.setBorderTop(BorderStyle.THICK);
    header.setBorderLeft(BorderStyle.THICK);
    header.setBorderRight(BorderStyle.THICK);
    header.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
    header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    header.setVerticalAlignment(VerticalAlignment.CENTER);
    header.setAlignment(HorizontalAlignment.CENTER);
    styles.put("header", header);

    XSSFCellStyle notExist = wb.createCellStyle();
    XSSFFont notExistFont = wb.createFont();
    notExistFont.setColor(HSSFColor.HSSFColorPredefined.ORANGE.getIndex());
    notExistFont.setFontName("Calibri");
    notExist.setFont(notExistFont);
    notExist.setBorderBottom(BorderStyle.HAIR);
    notExist.setBorderTop(BorderStyle.HAIR);
    notExist.setBorderLeft(BorderStyle.HAIR);
    notExist.setBorderRight(BorderStyle.HAIR);
    notExist.setFillForegroundColor(IndexedColors.RED.getIndex());
    notExist.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("notExist", notExist);

    return styles;
  }

  /**
   *
   * 
   * @throws Exception
   */
  public void generate(Map<String, XSSFCellStyle> styles, SharedStringsTable sst, 
      String key, List<List<ExportCell>> exportRows) throws Exception {
    CTRst st;
    beginXml();
    columnsWidth(columnsWidthMap.get(key));
    beginSheet();
    for(int rowNumber = 0 ; rowNumber < exportRows.size() ; rowNumber++) {
      insertRow(rowNumber);
      for(int columnNumber = 0 ; columnNumber < exportRows.get(rowNumber).size() ; columnNumber++) {
        if(exportRows.get(rowNumber).get(columnNumber) != null) {
          st = CTRst.Factory.newInstance();
          st.setT(exportRows.get(rowNumber).get(columnNumber).getText());
          if(exportRows.get(rowNumber).get(columnNumber).getStyle() != null) {
            createCell(columnNumber, sst.addEntry( st ),styles.get(exportRows.get(rowNumber).get(columnNumber).getStyle()).getIndex(),"s");
          } else {
            createCell(columnNumber, sst.addEntry( st ),-1,"s");
          }
        }
      }
      endRow();
    }
    endSheetData();
    mergeCell(mergeCellsMap.get(key));
    endSheet();
  }

  public void beginXml() throws IOException {
    sw.beginXml();
  }
  
  public void beginSheet() throws IOException {
    sw.beginSheet();
  }

  public void insertRow(int rowNum) throws IOException {
    sw.insertRow(rowNum);
  }

  public void createCell(int columnIndex, String value) throws IOException {
    sw.createCell(columnIndex, value, -1);
  }

  public void createCell(int columnIndex, String value, int styleIndex)
      throws IOException {
    sw.createCell(columnIndex, value, styleIndex);
  }

  public void createCell(int columnIndex, int value, int styleIndex,
      String Datatype) throws IOException {
    sw.createCell(columnIndex, value, styleIndex, Datatype);
  }

  public void createCell(int columnIndex, double value) throws IOException {
    sw.createCell(columnIndex, value, -1);
  }

  public void createCell(int columnIndex, double value, int styleIndex)
      throws IOException {
    sw.createCell(columnIndex, value, styleIndex);
  }

  public void createCell(int columnIndex, Calendar value, int styleIndex)
      throws IOException {
    createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
  }

  public void mergeCell(List<MergeCell> mergeCells)
      throws IOException {
    sw.mergeCell(mergeCells);
  }
  
  public void columnsWidth(List<Integer> columnsWidth)
      throws IOException {
    sw.columnsWidth(columnsWidth);
  }
  
  public void endRow() throws IOException {
    sw.endRow();
  }

  public void endSheetData() throws IOException {
    sw.endSheetData();
  }
  
  public void endSheet() throws IOException {
    sw.endSheet();
  }

  /**
   * 
   * @param zipfile
   *          the template file
   * @param tmpfile
   *          the XML file with the sheet data
   * @param entry
   *          the name of the sheet entry to substitute, e.g.
   *          xl/worksheets/sheet1.xml
   * @param out
   *          the stream to write the result to
   */
  private static void substitute(File zipfile, List<File> fileArray, List<String> sheetArray, 
      OutputStream out) throws IOException {
    ZipFile zip = new ZipFile(zipfile);
    ZipOutputStream zos = new ZipOutputStream(out);

    @SuppressWarnings("unchecked")
    Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
    while (en.hasMoreElements()) {
      ZipEntry ze = en.nextElement();
      if (!sheetArray.contains("/" + ze.getName())) {
        zos.putNextEntry(new ZipEntry(ze.getName()));
        InputStream is = zip.getInputStream(ze);
        copyStream(is, zos);
        is.close();
      }
    }
    
    for(int index = 0 ; index < fileArray.size() ; index++) {
      zos.putNextEntry(new ZipEntry(sheetArray.get(index).substring(1)));
      InputStream is = new FileInputStream(fileArray.get(index));
      copyStream(is, zos);
      is.close();
    }
    zos.close();
  }

  private static void copyStream(InputStream in, OutputStream out)
      throws IOException {
    byte[] chunk = new byte[1024];
    int count;
    while ((count = in.read(chunk)) >= 0) {
      out.write(chunk, 0, count);
    }
  }

  
  public static class SpreadsheetWriter {
    private final Writer _out;
    private int _rownum;
    private static String LINE_SEPARATOR = System.getProperty("line.separator");

    public SpreadsheetWriter(Writer out) {
      _out = out;
    }

    public void beginXml() throws IOException {
      _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
          + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
      /*_out.write("<sheetPr><pageSetUpPr fitToPage=\"true\"/></sheetPr>");
      _out.write("<dimension ref=\"A1\"/>");
      _out.write(
          "<sheetViews><sheetView workbookViewId=\"0\" tabSelected=\"true\"><pane ySplit=\"1.0\" state=\"frozen\" topLeftCell=\"A2\" activePane=\"bottomLeft\"/><selection pane=\"bottomLeft\"/></sheetView></sheetViews>");
      _out.write("<sheetFormatPr defaultRowHeight=\"15.0\"/>");*/
    }
    
    public void beginSheet() throws IOException {
      _out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void endSheetData() throws IOException {
      _out.write("</sheetData>");
    }
    
    public void endSheet() throws IOException {
      _out.write("</worksheet>");
    }

    /**
     * 
     * 
     * @param rownum
     * 
     */
    public void insertRow(int rownum) throws IOException {
      _out.write("<row r=\"" + (rownum + 1) + "\"");
      _out.write(">" + LINE_SEPARATOR);
      this._rownum = rownum;
    }

   
    public void endRow() throws IOException {
      _out.write("</row>" + LINE_SEPARATOR);
    }

    /**
     * 
     * 
     * @param columnIndex
     * @param value
     * @param styleIndex
     * @throws IOException
     */
    public void createCell(int columnIndex, String value, int styleIndex)
        throws IOException {
      String ref = new CellReference(_rownum, columnIndex).formatAsString();
      _out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
      if (styleIndex != -1)
        _out.write(" s=\"" + styleIndex + "\"");
      _out.write(">");
      _out.write("<is><t>" + encoderXML(value) + "</t></is>");
      _out.write("</c>");
    }

    public void createCell(int columnIndex, int value, int styleIndex,
        String Datatype) throws IOException {
      String ref = new CellReference(_rownum, columnIndex).formatAsString();
      _out.write("<c r=\"" + ref + "\" t=\"" + Datatype + "\"");
      if (styleIndex != -1)
        _out.write(" s=\"" + styleIndex + "\"");
      _out.write(">");
      _out.write("<v>" + value + "</v>");
      _out.write("</c>");
    }

    /*
     * public void createCell(int columnIndex, String value) throws IOException
     * { createCell(columnIndex, value, -1); }
     */

    public void createCell(int columnIndex, double value, int styleIndex)
        throws IOException {
      String ref = new CellReference(_rownum, columnIndex).formatAsString();
      _out.write("<c r=\"" + ref + "\"");
      if (styleIndex != -1)
        _out.write(" s=\"" + styleIndex + "\"");
      _out.write(" t=\"n\"");
      _out.write(">");
      _out.write("<v>" + value + "</v>");
      _out.write("</c>");
    }

    public void mergeCell(List<MergeCell> mergeCells)
        throws IOException {
      if(mergeCells != null) {
        _out.write("<mergeCells count=\""+mergeCells.size()+"\">");
        for(MergeCell mergeCell : mergeCells)
          _out.write("<mergeCell ref=\""+mergeCell.getStartCell()+":"+mergeCell.getEndCell()+"\"/>");
        _out.write("</mergeCells>");
      }
    }
    
    public void columnsWidth( List<Integer> columnsWidth)
        throws IOException {
      _out.write("<cols>");
      for(int index = 0 ; index < columnsWidth.size() ; index++) {
        _out.write("<col min=\""+(index+1)+"\" max=\""+(index+1)+"\" width=\""+columnsWidth.get(index)+"\" customWidth=\""+(index+1)+"\"/>");
      }
      _out.write("</cols>");
    }
    /*
     * public void createCell(int columnIndex, double value) throws IOException
     * { createCell(columnIndex, value, -1); }
     */

    /*
     * public void createCell(int columnIndex, Calendar value, int styleIndex)
     * throws IOException { createCell(columnIndex, DateUtil.getExcelDate(value,
     * false), styleIndex); }
     */
  }

  // XML Encode
  private static final String[] xmlCode = new String[256];

  static {
    // Special characters
    xmlCode['\''] = "'";
    xmlCode['\"'] = "\""; // double quote
    xmlCode['&'] = "&"; // ampersand
    xmlCode['<'] = "<"; // lower than
    xmlCode['>'] = ">"; // greater than
  }

  /**
   * <p>
   * Encode the given text into xml.
   * </p>
   * 
   * @param string
   *          the text to encode
   * @return the encoded string
   */
  public static String encoderXML(String string) {
    if (string == null)
      return "";
    int n = string.length();
    char character;
    String xmlchar;
    StringBuffer buffer = new StringBuffer();
    // loop over all the characters of the String.
    for (int i = 0; i < n; i++) {
      character = string.charAt(i);
      // the xmlcode of these characters are added to a StringBuffer
      // one by one
      try {
        xmlchar = xmlCode[character];
        if (xmlchar == null) {
          buffer.append(character);
        } else {
          buffer.append(xmlCode[character]);
        }
      } catch (ArrayIndexOutOfBoundsException aioobe) {
        buffer.append(character);
      }
    }
    return buffer.toString();
  }

  /**
   * 
   */
  public static void main(String[] args) throws Exception {

    String file = "C:\\Users\\wilson82740\\Desktop\\test5.xlsx";

    ExcelWriter writer = new ExcelWriter() {
      public void generate(Map<String, XSSFCellStyle> styles,
          SharedStringsTable sst, PackagePart part) throws Exception {
        Random rnd = new Random();
        Calendar calendar = Calendar.getInstance();
        this.beginSheet();

        for (int rownum = 0; rownum < 100; rownum++) {

          if (rownum == 0) {
            this.insertRow(0);
            int styleIndex = styles.get("header").getIndex();
            this.createCell(0, "Title", styleIndex);
            this.createCell(1, "% Change", styleIndex);
            this.createCell(2, "Ratio", styleIndex);
            this.createCell(3, "Expenses", styleIndex);
            this.createCell(4, "Date", styleIndex);

            this.endRow();
          } else {
            this.insertRow(rownum);

            this.createCell(0, "Hellodsfgdfsgsdfgsdfgsdfgsdfgsdfgsgsdfgsd, " + rownum + "!", styles.get("data").getIndex());
            this.createCell(1, (double) rnd.nextInt(100) / 100,
                styles.get("percent").getIndex());
            this.createCell(2, (double) rnd.nextInt(10) / 10,
                styles.get("coeff").getIndex());
            /*this.createCell(3, rnd.nextInt(10000),
                styles.get("currency").getIndex());*/
            this.createCell(4, calendar, styles.get("date").getIndex());

            this.endRow();

            calendar.roll(Calendar.DAY_OF_YEAR, 4);
          }
        }
        
        this.insertRow(100);
        this.createCell(0, "Hello, 100!", styles.get("data").getIndex());
        this.endRow();
        this.insertRow(101);
        this.createCell(0, "", styles.get("data").getIndex());
        this.endRow();
        
        this.endSheetData();
        //this.mergeCell();
        this.endSheet();
      }
    };
    writer.process(file);
  }

}