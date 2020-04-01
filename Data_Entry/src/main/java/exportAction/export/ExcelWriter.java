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
      XSSFSheet sheet = wb.createSheet(entry.getKey());
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

    XSSFCellStyle style1 = wb.createCellStyle();
    style1.setAlignment(HorizontalAlignment.RIGHT);
    style1.setDataFormat(fmt.getFormat("0.0%"));
    styles.put("percent", style1);

    XSSFCellStyle style2 = wb.createCellStyle();
    style2.setAlignment(HorizontalAlignment.CENTER);
    style2.setDataFormat(fmt.getFormat("0.0X"));
    styles.put("coeff", style2);

    XSSFCellStyle style3 = wb.createCellStyle();
    style3.setAlignment(HorizontalAlignment.RIGHT);
    styles.put("RightAlignment", style3);

    XSSFCellStyle style4 = wb.createCellStyle();
    style4.setAlignment(HorizontalAlignment.RIGHT);
    style4.setShrinkToFit(true);
    style4.setDataFormat(fmt.getFormat("yyyy/MM/dd"));
    styles.put("date", style4);

    XSSFCellStyle style5 = wb.createCellStyle();
    XSSFFont headerFont = wb.createFont();
    headerFont.setBold(true);
    style5.setBorderBottom(BorderStyle.HAIR);
    style5.setBorderTop(BorderStyle.HAIR);
    style5.setBorderLeft(BorderStyle.HAIR);
    style5.setBorderRight(BorderStyle.HAIR);
    style5.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
    style5.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    style5.setFont(headerFont);
    style5.setVerticalAlignment(VerticalAlignment.CENTER);
    style5.setAlignment(HorizontalAlignment.CENTER);
    styles.put("header", style5);

    XSSFCellStyle style6 = wb.createCellStyle();
    XSSFFont BorderFont = wb.createFont();
    BorderFont.setFontName("Calibri");
    BorderFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
    BorderFont.setBold(true);
    style6.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
    style6.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
    style6.setBorderBottom(BorderStyle.MEDIUM);
    style6.setBorderTop(BorderStyle.MEDIUM);
    style6.setFont(BorderFont);
    styles.put("Border", style6);
    
    XSSFCellStyle style7 = wb.createCellStyle();
    style7.setBorderBottom(BorderStyle.HAIR);
    style7.setBorderTop(BorderStyle.HAIR);
    style7.setBorderLeft(BorderStyle.HAIR);
    style7.setBorderRight(BorderStyle.HAIR);
    style7.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
    style7.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("inconsistent", style7);
    
    XSSFCellStyle style8 = wb.createCellStyle();
    XSSFFont notFoundFont = wb.createFont();
    notFoundFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
    style8.setFont(notFoundFont);
    styles.put("notFound", style8);
    
    XSSFCellStyle style9 = wb.createCellStyle();
    style9.setVerticalAlignment(VerticalAlignment.CENTER);
    style9.setAlignment(HorizontalAlignment.CENTER);
    styles.put("merge", style9);
    
    XSSFCellStyle style10 = wb.createCellStyle();
    XSSFFont notExist = wb.createFont();
    notExist.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
    style10.setFont(notExist);
    style10.setBorderBottom(BorderStyle.HAIR);
    style10.setBorderTop(BorderStyle.HAIR);
    style10.setBorderLeft(BorderStyle.HAIR);
    style10.setBorderRight(BorderStyle.HAIR);
    style10.setFillForegroundColor(IndexedColors.RED.getIndex());
    style10.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styles.put("notExist", style10);

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