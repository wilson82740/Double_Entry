package exportAction;

import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;

import org.apache.ibatis.session.SqlSession;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import com.avaje.ebean.Ebean;

import exportAction.export.ExcelWriter;
import dao.PatientMapper;
import dao.Same_MRNMapper;
import exporModel.ColumnCell;
import exporModel.ExportCell;
import exporModel.MergeCell;
import model.CrfDatarecordBean;
import model.DataRecordBean;
import model.Patient;
import model.PatientDatapoint;
import model.PatientExample;
import model.PatientVisit;
import model.PatienttoMRNBean;
import model.Same_MRN;

public class ExportAction {

  public static final String notFound = " is not found in ";
  public static final String inconsistent = " is inconsistent with other protocol. ";
  public static final String bothInconsistent = " are both inconsistent with other protocol. ";
  public static final String formNotExist = " This form doesn't exist. ";
  public static final String diffanswer = " Same question but different answer. Please reconfirm. ";
  public static final String fieldEmpty = " The field is empty in ";
  
  public static void exportLogPage(ExcelWriter writer, String protocolnum1, String protocolnum2) {
    DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
    LocalDateTime now = LocalDateTime.now();
    List<List<ExportCell>> exportRows = new ArrayList<List<ExportCell>>();
    List<ExportCell> exportRow = null;
    
    exportRow = new ArrayList<ExportCell>();
    ExportCell cell = null;
    cell = new ExportCell();
    cell.setText("Date");
    cell.setStyle("header");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText(dtf.format(now));
    cell.setStyle("normal");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText("Number of Protocol");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Website");
    cell.setStyle("header");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText(protocolnum1 + "_1");
    cell.setStyle("normal");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("https://csis.cims.tw/med");
    cell.setStyle("normal");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText(protocolnum2 + "_2");
    cell.setStyle("normal");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("https://csis.cims.tw/med2");
    cell.setStyle("normal");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    List<ColumnCell> columns = new ArrayList<ColumnCell>();
    ColumnCell column ;
    column = new ColumnCell();
    column.setWidth(25);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    
    try {
      if(writer.getExportRowsMap() == null) {
        Map<String, List<List<ExportCell>>> exportRowsMap = new LinkedHashMap<String, List<List<ExportCell>>>();
        exportRowsMap.put("Log Page", exportRows);
        writer.setExportRowsMap(exportRowsMap);
        Map<String, List<ColumnCell>> columnsMap = new LinkedHashMap<String, List<ColumnCell>>();
        columnsMap.put("Log Page", columns);
        writer.setColumnsMap(columnsMap);
        Map<String, List<MergeCell>> mergeCellsMap = new LinkedHashMap<String, List<MergeCell>>();
        mergeCellsMap.put("Log Page", null);
        writer.setMergeCellsMap(mergeCellsMap);
      } else {
        writer.getExportRowsMap().put("Log Page", exportRows);
        writer.getColumnsMap().put("Log Page", columns);
        writer.getMergeCellsMap().put("Log Page", null);
      }
      //writer.process("C:\\Users\\wilson82740\\Desktop\\Patient.xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
  
  public static void exportpatient(ExcelWriter writer, long protocol1Id, long protocol2Id, String protocolnum1, 
      String protocolnum2, List<PatienttoMRNBean> patienttoMRNs, PatientMapper patientMapperMed1, PatientMapper patientMapperMed2) {
    List<MergeCell> mergeCells = new ArrayList<MergeCell>();
    List<List<ExportCell>> exportRows = new ArrayList<List<ExportCell>>();
    List<ExportCell> exportRow = null;

    MergeCell mergeCell = null;
    mergeCell = new MergeCell();
    mergeCell.setStartCell(new CellReference(0, 1).formatAsString());
    mergeCell.setEndCell(new CellReference(0, 2).formatAsString());
    mergeCells.add(mergeCell);
    mergeCell = new MergeCell();
    mergeCell.setStartCell(new CellReference(0, 3).formatAsString());
    mergeCell.setEndCell(new CellReference(0, 4).formatAsString());
    mergeCells.add(mergeCell);
    mergeCell = new MergeCell();
    mergeCell.setStartCell(new CellReference(0, 5).formatAsString());
    mergeCell.setEndCell(new CellReference(1, 5).formatAsString());
    mergeCells.add(mergeCell);
    
    exportRow = new ArrayList<ExportCell>();
    ExportCell cell = null;

    exportRow.add(null);
    cell = new ExportCell();
    cell.setText(protocolnum1 + "_1");
    cell.setStyle("protocolTopic1");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("");
    cell.setStyle("protocolTopic1");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText(protocolnum2 + "_2");
    cell.setStyle("protocolTopic2");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("");
    cell.setStyle("protocolTopic2");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Inconsistent Type");
    cell.setStyle("inconsistentTopic");
    exportRow.add(cell);
    exportRows.add(exportRow);

    exportRow = new ArrayList<ExportCell>();

    cell = new ExportCell();
    cell.setText("Subject ID");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Last Name");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("GUID");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Last Name");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("GUID");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("");
    cell.setStyle("inconsistentTopic");
    exportRow.add(cell);
    exportRows.add(exportRow);

    for (int index = 0; index < patienttoMRNs.size(); index++) {
      if (patienttoMRNs.get(index).getPatient2Id() == null && patienttoMRNs.get(index).getPatient1Id() != null) {
        exportRow = new ArrayList<ExportCell>();
        PatientExample patientexample1 = new PatientExample();
        patientexample1.or().andPatientidEqualTo(patienttoMRNs.get(index).getPatient1Id())
            .andProtocolidEqualTo(protocol1Id);
        Patient patient1 =
            patientMapperMed1.noStudyEventByPrimaryKey(patientexample1);
        cell = new ExportCell();
        cell.setText(patienttoMRNs.get(index).getMrn());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patient1.getLastName());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patient1.getFirstName());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText("");
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText("");
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patienttoMRNs.get(index).getMrn() + notFound + protocolnum2 + "_2" + ".");
        cell.setStyle("bold");
        exportRow.add(cell);
        exportRows.add(exportRow);
      } else if(patienttoMRNs.get(index).getPatient2Id() != null && patienttoMRNs.get(index).getPatient1Id() == null) {
        exportRow = new ArrayList<ExportCell>();
        PatientExample patientexample2 = new PatientExample();
        patientexample2.or().andPatientidEqualTo(patienttoMRNs.get(index).getPatient2Id())
        .andProtocolidEqualTo(protocol2Id);
        Patient patient2 =
            patientMapperMed2.noStudyEventByPrimaryKey(patientexample2);
        cell = new ExportCell();
        cell.setText(patient2.getMrn());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText("");
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText("");
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patient2.getLastName());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patient2.getFirstName());
        cell.setStyle("normal");
        exportRow.add(cell);
        cell = new ExportCell();
        cell.setText(patient2.getMrn() + notFound + protocolnum1 + "_1" + ".");
        cell.setStyle("bold");
        exportRow.add(cell);
        exportRows.add(exportRow);
      } else {
        PatientExample patientexample1 = new PatientExample();
        patientexample1.or().andPatientidEqualTo(patienttoMRNs.get(index).getPatient1Id())
        .andProtocolidEqualTo(protocol1Id);
        Patient patient1 =
            patientMapperMed1.noStudyEventByPrimaryKey(patientexample1);
        PatientExample patientexample2 = new PatientExample();
        patientexample2.or().andPatientidEqualTo(patienttoMRNs.get(index).getPatient2Id())
        .andProtocolidEqualTo(protocol2Id);
        Patient patient2 =
            patientMapperMed2.noStudyEventByPrimaryKey(patientexample2);
        if (!patienttoMRNs.get(index).getSameLastName() && patienttoMRNs.get(index).getSameFirstName()) {
          exportRow = new ArrayList<ExportCell>();
          cell = new ExportCell();
          cell.setText(patienttoMRNs.get(index).getMrn());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText("The last name of "+ patienttoMRNs.get(index).getMrn() + inconsistent);
          cell.setStyle("bold");
          exportRow.add(cell);
          exportRows.add(exportRow);
        } else if (patienttoMRNs.get(index).getSameLastName() && !patienttoMRNs.get(index).getSameFirstName()) {
          exportRow = new ArrayList<ExportCell>();
          cell = new ExportCell();
          cell.setText(patienttoMRNs.get(index).getMrn());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText("The GUID of "+ patienttoMRNs.get(index).getMrn() + inconsistent);
          cell.setStyle("bold");
          exportRow.add(cell);
          exportRows.add(exportRow);
        } else if(!patienttoMRNs.get(index).getSameLastName() && !patienttoMRNs.get(index).getSameFirstName()) {
          exportRow = new ArrayList<ExportCell>();
          cell = new ExportCell();
          cell.setText(patienttoMRNs.get(index).getMrn());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient1.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getLastName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText(patient2.getFirstName());
          cell.setStyle("normal");
          exportRow.add(cell);
          cell = new ExportCell();
          cell.setText("The GUID & last name of "+ patienttoMRNs.get(index).getMrn() + bothInconsistent);
          cell.setStyle("bold");
          exportRow.add(cell);
          exportRows.add(exportRow);
        } else {
          Same_MRN sameMRN = new Same_MRN();
          sameMRN.setMrn(patienttoMRNs.get(index).getMrn());
          sameMRN.setPatient1Id(patienttoMRNs.get(index).getPatient1Id());
          sameMRN.setPatient2Id(patienttoMRNs.get(index).getPatient2Id());
          Ebean.save(sameMRN);
        }
      }
    }
    
    List<ColumnCell> columns = new ArrayList<ColumnCell>();
    ColumnCell column ;
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(50);
    column.setHide(0);
    columns.add(column);

    try {
      if(writer.getExportRowsMap() == null) {
        Map<String, List<List<ExportCell>>> exportRowsMap = new LinkedHashMap<String, List<List<ExportCell>>>();
        exportRowsMap.put("Patient", exportRows);
        writer.setExportRowsMap(exportRowsMap);
        Map<String, List<ColumnCell>> columnsMap = new LinkedHashMap<String, List<ColumnCell>>();
        columnsMap.put("Patient", columns);
        writer.setColumnsMap(columnsMap);
        Map<String, List<MergeCell>> mergeCellsMap = new LinkedHashMap<String, List<MergeCell>>();
        mergeCellsMap.put("Patient", mergeCells);
        writer.setMergeCellsMap(mergeCellsMap);
      } else {
        writer.getExportRowsMap().put("Patient", exportRows);
        writer.getColumnsMap().put("Patient", columns);
        writer.getMergeCellsMap().put("Patient", mergeCells);
      }
      //writer.process("C:\\Users\\wilson82740\\Desktop\\Patient.xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }

  public static void exportVisit(ExcelWriter writer, String protocolnum1, String protocolnum2, List<PatientVisit> diffVisitDatePatients, List<String> allVisit) {
    
    String cellText = null;
    SimpleDateFormat sdFormat = new SimpleDateFormat("yyyy/MM/dd");
    
    List<List<ExportCell>> exportRows = new ArrayList<List<ExportCell>>();
    List<MergeCell> mergeCells = null;
    if(!diffVisitDatePatients.isEmpty())
        mergeCells = new ArrayList<MergeCell>();
    List<ExportCell> exportRow = null;
    MergeCell mergeCell = null;

    exportRow = new ArrayList<ExportCell>();
    ExportCell cell = null;
    
    cell = new ExportCell();
    cell.setText("Subject ID / Visit");
    cell.setStyle("boldandBorder");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Protocol Name");
    cell.setStyle("boldandBorder");
    exportRow.add(cell);
    if(allVisit != null) {
      for(String visitName : allVisit) {
        cell = new ExportCell();
        cell.setText(visitName);
        cell.setStyle("boldandBorder");
        exportRow.add(cell);
      }
    }
    exportRows.add(exportRow);
    
    Map<String,Integer> mrnLocationMap = new HashMap<String,Integer>();
    
    for(PatientVisit diffVisitDatePatient : diffVisitDatePatients) {
      Integer mrnLocation = mrnLocationMap.get(diffVisitDatePatient.getMrn());
      if(mrnLocation == null) {
        mrnLocationMap.put(diffVisitDatePatient.getMrn(), exportRows.size());
        mergeCell = new MergeCell();
        mergeCell.setStartCell(new CellReference(exportRows.size(), 0).formatAsString());
        mergeCell.setEndCell(new CellReference(exportRows.size()+1, 0).formatAsString());
        mergeCells.add(mergeCell);
        List<ExportCell> exportRow1 = new ArrayList<ExportCell>();
        List<ExportCell> exportRow2 = new ArrayList<ExportCell>();
        cell = new ExportCell();
        cell.setText(diffVisitDatePatient.getMrn());
        cell.setStyle("boldandBorder");
        exportRow1.add(cell);
        cell = new ExportCell();
        cell.setText(protocolnum1 + "_1");
        cell.setStyle("tlBorder");
        exportRow1.add(cell);
        cell = new ExportCell();
        cell.setText("");
        cell.setStyle("boldandBorder");
        exportRow2.add(cell);
        cell = new ExportCell();
        cell.setText(protocolnum2 + "_2");
        cell.setStyle("blBorder");
        exportRow2.add(cell);
        for(int index = 0 ; index < allVisit.size() ; index++) {
          if(index == allVisit.size()-1) {
            cell = new ExportCell();
            cell.setText("");
            cell.setStyle("trBorder");
            exportRow1.add(cell);
            cell = new ExportCell();
            cell.setText("");
            cell.setStyle("brBorder");
            exportRow2.add(cell);
          } else {
            cell = new ExportCell();
            cell.setText("");
            cell.setStyle("topBorder");
            exportRow1.add(cell);
            cell = new ExportCell();
            cell.setText("");
            cell.setStyle("bottomBorder");
            exportRow2.add(cell);
          }
        }
        cell = new ExportCell(); 
        cellText = diffVisitDatePatient.getVisitDate1() == null ? "EMPTY" : sdFormat.format(diffVisitDatePatient.getVisitDate1());
        cell.setText(cellText);
        if(diffVisitDatePatient.getVisitDate1() == null || diffVisitDatePatient.getVisitDate2() == null) {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("greentrBorder");
          } else {
            cell.setStyle("greenTopBorder");
          }
        } else {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("yellowtrBorder");
          } else {
            cell.setStyle("yellowTopBorder");
          }
        }
        exportRow1.set(diffVisitDatePatient.getEventOrder().intValue()+1, cell);
        cell = new ExportCell();
        cellText = diffVisitDatePatient.getVisitDate2() == null ? "EMPTY" : sdFormat.format(diffVisitDatePatient.getVisitDate2());
        cell.setText(cellText);
        if(diffVisitDatePatient.getVisitDate1() == null || diffVisitDatePatient.getVisitDate2() == null) {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("greenbrBorder");
          } else {
            cell.setStyle("greenBottomBorder");
          }
        } else {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("yellowbrBorder");
          } else {
            cell.setStyle("yellowBottomBorder");
          }
        }
        exportRow2.set(diffVisitDatePatient.getEventOrder().intValue()+1, cell);
        exportRows.add(exportRow1);
        exportRows.add(exportRow2);
      } else {
        List<ExportCell> exportRow1 = exportRows.get(mrnLocation);
        List<ExportCell> exportRow2 = exportRows.get(mrnLocation+1);
        cell = new ExportCell(); 
        cellText = diffVisitDatePatient.getVisitDate1() == null ? "EMPTY" : sdFormat.format(diffVisitDatePatient.getVisitDate1());
        cell.setText(cellText);
        if(diffVisitDatePatient.getVisitDate1() == null || diffVisitDatePatient.getVisitDate2() == null) {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("greentrBorder");
          } else {
            cell.setStyle("greenTopBorder");
          }
        } else {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("yellowtrBorder");
          } else {
            cell.setStyle("yellowTopBorder");
          }
        }
        exportRow1.set(diffVisitDatePatient.getEventOrder().intValue()+1, cell);
        cell = new ExportCell();
        cellText = diffVisitDatePatient.getVisitDate2() == null ? "EMPTY" : sdFormat.format(diffVisitDatePatient.getVisitDate2());
        cell.setText(cellText);
        if(diffVisitDatePatient.getVisitDate1() == null || diffVisitDatePatient.getVisitDate2() == null) {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("greenbrBorder");
          } else {
            cell.setStyle("greenBottomBorder");
          }
        } else {
          if(diffVisitDatePatient.getEventOrder().intValue() == allVisit.size()) {
            cell.setStyle("yellowbrBorder");
          } else {
            cell.setStyle("yellowBottomBorder");
          }
        }
        exportRow2.set(diffVisitDatePatient.getEventOrder().intValue()+1, cell);
      }
    }
    
    List<ColumnCell> columns = new ArrayList<ColumnCell>();
    ColumnCell column ;
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(30);
    column.setHide(0);
    columns.add(column);
    for(int visitIndex = 0 ; visitIndex < exportRows.get(0).size()-2 ; visitIndex++) {
      column = new ColumnCell();
      column.setWidth(20);
      column.setHide(0);
      columns.add(column);
    }
    
    try {
      if(writer.getExportRowsMap() == null) {
        Map<String, List<List<ExportCell>>> exportRowsMap = new LinkedHashMap<String, List<List<ExportCell>>>();
        exportRowsMap.put("Visit_Date", exportRows);
        writer.setExportRowsMap(exportRowsMap);
        Map<String, List<ColumnCell>> columnsMap = new LinkedHashMap<String, List<ColumnCell>>();
        columnsMap.put("Visit_Date", columns);
        writer.setColumnsMap(columnsMap);
        Map<String, List<MergeCell>> mergeCellsMap = new LinkedHashMap<String, List<MergeCell>>();
        mergeCellsMap.put("Visit_Date", mergeCells);
        writer.setMergeCellsMap(mergeCellsMap);
      } else {
        Map<String, List<List<ExportCell>>> exportRowsMap = writer.getExportRowsMap();
        exportRowsMap.put("Visit_Date", exportRows);
        writer.getColumnsMap().put("Visit_Date", columns);
        writer.getMergeCellsMap().put("Visit_Date", mergeCells);
      }
      //writer.process("C:\\Users\\wilson82740\\Desktop\\Visit_Date.xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
  
  public static void exportCRFStauts(ExcelWriter writer, List<CrfDatarecordBean> exportCRFStauts, Map<String, Map<String,Integer>> allCRFinVisit, String protocolnum1, String protocolnum2) {
    List<List<ExportCell>> exportRows = new ArrayList<List<ExportCell>>();
    List<ExportCell> exportRow = null;
    ExportCell cell = null;
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText("Subject ID");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Visit Name");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Form Name .ver (#)");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Page");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Form Status in "+protocolnum1 + "_1");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Form Status in "+protocolnum2 + "_2");
    cell.setStyle("header");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    for(CrfDatarecordBean stauts : exportCRFStauts) {
      exportRow = new ArrayList<ExportCell>();
      cell = new ExportCell();
      cell.setText(stauts.getMrn());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(stauts.getVisitName());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      if(stauts.getEventSequence() == 1 ) {
        cell.setText(stauts.getTitle());
      } else {
        cell.setText(stauts.getTitle()+" | #"+stauts.getEventSequence());
      }
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(stauts.getPage());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(stauts.getStauts1());
      if(stauts.getStauts1().equals(formNotExist)) {
        cell.setStyle("notExist");
      } else {
        cell.setStyle("normal");
      }
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(stauts.getStauts2());
      if(stauts.getStauts2().equals(formNotExist)) {
        cell.setStyle("notExist");
      } else {
        cell.setStyle("normal");
      }
      exportRow.add(cell);
      exportRows.add(exportRow);
    }
    
    List<ColumnCell> columns = new ArrayList<ColumnCell>();
    ColumnCell column ;
    column = new ColumnCell();
    column.setWidth(20);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(15);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(15);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    
    try {
      if(writer.getExportRowsMap() == null) {
        Map<String, List<List<ExportCell>>> exportRowsMap = new LinkedHashMap<String, List<List<ExportCell>>>();
        exportRowsMap.put("Form Stauts", exportRows);
        writer.setExportRowsMap(exportRowsMap);
        Map<String, List<ColumnCell>> columnsMap = new LinkedHashMap<String, List<ColumnCell>>();
        columnsMap.put("Form Stauts", columns);
        writer.setColumnsMap(columnsMap);
        Map<String, List<MergeCell>> mergeCellsMap = new LinkedHashMap<String, List<MergeCell>>();
        mergeCellsMap.put("Form Stauts", null);
        writer.setMergeCellsMap(mergeCellsMap);
      } else {
        writer.getExportRowsMap().put("Form Stauts", exportRows);
        writer.getColumnsMap().put("Form Stauts", columns);
        writer.getMergeCellsMap().put("Form Stauts", null);
      }
      //writer.process("C:\\Users\\wilson82740\\Desktop\\Visit_Form.xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
  public static void exportDatapoint(ExcelWriter writer, List<PatientDatapoint> exportDatapoints, String protocolnum1, 
      String protocolnum2, String sheetName, Boolean hasPageColumn) {
    List<List<ExportCell>> exportRows = new ArrayList<List<ExportCell>>();
    List<ExportCell> exportRow = null;
    ExportCell cell = null;
    
    int valueEmpty;
    
    exportRow = new ArrayList<ExportCell>();
    cell = new ExportCell();
    cell.setText("Subject ID");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Form Name .ver (#)");
    cell.setStyle("header");
    exportRow.add(cell);
    if(hasPageColumn != null) {
      cell = new ExportCell();
      cell.setText("Page");
      cell.setStyle("header");
      exportRow.add(cell);
    }
    cell = new ExportCell();
    cell.setText("Question Text");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Variable Name");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Value in "+protocolnum1 + "_1");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Value in "+protocolnum2 + "_2");
    cell.setStyle("header");
    exportRow.add(cell);
    cell = new ExportCell();
    cell.setText("Query Message");
    cell.setStyle("header");
    exportRow.add(cell);
    exportRows.add(exportRow);
    
    for(PatientDatapoint exportDatapoint : exportDatapoints) {
      valueEmpty = -1;
      exportRow = new ArrayList<ExportCell>();
      cell = new ExportCell();
      cell.setText(exportDatapoint.getMrn());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(exportDatapoint.getFormTitle());
      cell.setStyle("normal");
      exportRow.add(cell);
      if(hasPageColumn != null) {
        cell = new ExportCell();
        cell.setText(exportDatapoint.getPage());
        cell.setStyle("normal");
        exportRow.add(cell);
      }
      cell = new ExportCell();
      cell.setText(exportDatapoint.getQuestionText());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      cell.setText(exportDatapoint.getDatapointName());
      cell.setStyle("normal");
      exportRow.add(cell);
      cell = new ExportCell();
      if(exportDatapoint.getValue1() == null) {
        valueEmpty = 1;
        cell.setText("EMPTY");
        cell.setStyle("valueEmpty");
      } else {
        cell.setText(exportDatapoint.getValue1());
        cell.setStyle("normal");
      }
      exportRow.add(cell);
      cell = new ExportCell();
      if(exportDatapoint.getValue2() == null) {
        valueEmpty = 2;
        cell.setText("EMPTY");
        cell.setStyle("valueEmpty");
      } else {
        cell.setText(exportDatapoint.getValue2());
        cell.setStyle("normal");
      }
      exportRow.add(cell);
      cell = new ExportCell();
      if(valueEmpty != -1) {
        if(valueEmpty == 1) {
          cell.setText(fieldEmpty + protocolnum1 + "_1" + ".");
        } else {
          cell.setText(fieldEmpty + protocolnum2 + "_2" + ".");
        }
      } else {
        cell.setText(diffanswer);
      }
      cell.setStyle("bold");
      exportRow.add(cell);
      exportRows.add(exportRow);
    }
    
    List<ColumnCell> columns = new ArrayList<ColumnCell>();
    ColumnCell column ;
    column = new ColumnCell();
    column.setWidth(20);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    if(hasPageColumn != null) {
      column = new ColumnCell();
      column.setWidth(15);
      column.setHide(0);
      columns.add(column);
    }
    column = new ColumnCell();
    column.setWidth(80);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(15);
    column.setHide(1);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(40);
    column.setHide(0);
    columns.add(column);
    column = new ColumnCell();
    column.setWidth(50);
    column.setHide(1);
    columns.add(column);
    
    try {
      if(writer.getExportRowsMap() == null) {
        Map<String, List<List<ExportCell>>> exportRowsMap = new LinkedHashMap<String, List<List<ExportCell>>>();
        exportRowsMap.put(sheetName, exportRows);
        writer.setExportRowsMap(exportRowsMap);
        Map<String, List<ColumnCell>> columnsMap = new LinkedHashMap<String, List<ColumnCell>>();
        columnsMap.put(sheetName, columns);
        writer.setColumnsMap(columnsMap);
        Map<String, List<MergeCell>> mergeCellsMap = new LinkedHashMap<String, List<MergeCell>>();
        mergeCellsMap.put(sheetName, null);
        writer.setMergeCellsMap(mergeCellsMap);
      } else {
        writer.getExportRowsMap().put(sheetName, exportRows);
        writer.getColumnsMap().put(sheetName, columns);
        writer.getMergeCellsMap().put(sheetName, null);
      }
      //writer.process("C:\\Users\\wilson82740\\Desktop\\Datapoint.xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
}
