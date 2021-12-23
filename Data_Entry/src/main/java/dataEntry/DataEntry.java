package dataEntry;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.ibatis.session.SqlSession;

import com.avaje.ebean.Ebean;

import dao.Crf_DatarecordMapper;
import dao.DatapointDatarecordMapper;
import dao.DatapointMapper;
import dao.Match_VisitMapper;
import dao.PatientMapper;
import dao.ProtocolMapper;
import dao.Same_MRNMapper;
import dao.VisitMapper;
import exportAction.ExportAction;
import exportAction.export.ExcelWriter;
import importAction.ExcelReader;
import model.CrfDatarecordBean;
import model.DataRecordBean;
import model.Datapoint;
import model.DatapointDatarecord;
import model.DatapointDatarecordExample;
import model.EnumType;
import model.Patient;
import model.PatientDatapoint;
import model.PatientExample;
import model.PatientVisit;
import model.PatienttoMRNBean;
import model.Protocol;
import model.Same_MRN;
import model.StudyEvent;
import model.ViewCRFBean;
import model.Visit;
import model.VisitExample;

public class DataEntry {
  
  public void runEntey(long protocol1Id, long protocol2Id) throws IOException{
    System.out.println("runEntey");
    MybatisConnect mybatisConnect = new MybatisConnect();
    mybatisConnect.setSqlSession("med1.xml", "med2.xml", "DataEntryMapperConfig.xml");
    ProtocolMapper protocolMapperMed1 = mybatisConnect.getSqlSessionMed1().getMapper(ProtocolMapper.class);
    PatientMapper patientMapperMed1 = mybatisConnect.getSqlSessionMed1().getMapper(PatientMapper.class);
    VisitMapper visitMapperMed1 = mybatisConnect.getSqlSessionMed1().getMapper(VisitMapper.class);
    DatapointMapper datapointMapperMed1 = mybatisConnect.getSqlSessionMed1().getMapper(DatapointMapper.class);
    DatapointDatarecordMapper datapointDatarecordMapperMed1 = mybatisConnect.getSqlSessionMed1().getMapper(DatapointDatarecordMapper.class);
    ProtocolMapper protocolMapperMed2 = mybatisConnect.getSqlSessionMed2().getMapper(ProtocolMapper.class);
    PatientMapper patientMapperMed2 = mybatisConnect.getSqlSessionMed2().getMapper(PatientMapper.class);
    VisitMapper visitMapperMed2 = mybatisConnect.getSqlSessionMed2().getMapper(VisitMapper.class);
    DatapointMapper datapointMapperMed2 = mybatisConnect.getSqlSessionMed2().getMapper(DatapointMapper.class);
    DatapointDatarecordMapper datapointDatarecordMapperMed2 = mybatisConnect.getSqlSessionMed2().getMapper(DatapointDatarecordMapper.class);
//    Same_MRNMapper sameMRNMapper = mybatisConnect.getSqlSessionDataEntry().getMapper(Same_MRNMapper.class);
//    Match_VisitMapper matchVisitMapper = mybatisConnect.getSqlSessionDataEntry().getMapper(Match_VisitMapper.class);
//    Crf_DatarecordMapper CDMapper = mybatisConnect.getSqlSessionDataEntry().getMapper(Crf_DatarecordMapper.class);
    
    Map<String, Boolean> hasPageColumnMap = new TreeMap<String, Boolean>();
    
    String folderName = "";
    File thisFile = null;
    if(protocol1Id == 1007) {
      thisFile= new File("Control.xlsx");
      folderName = "Control group";
    } else {
      thisFile= new File("Ricovir.xlsx");
      folderName = "Ricovir group";
    }
    
    List<List<List<String>>> excel = null;
    if(thisFile != null) {
      ExcelReader excelReader = new ExcelReader();
      try {
        excel = excelReader.processAll(Files.newInputStream(thisFile.toPath()));//透過xml讀取檔案資料
      } catch (Exception e) {
        System.out.println("Error:"+ e);
      }
    }
    
    ExcelWriter writer = new ExcelWriter() {};
    ExcelWriter writerDP = new ExcelWriter() {};

    //取得計畫名稱
    Protocol protocol1 = protocolMapperMed1.selectByPrimaryKey(protocol1Id);
    Protocol protocol2 = protocolMapperMed2.selectByPrimaryKey(protocol2Id);
    
    ExportAction.exportLogPage(writer, protocol1.getProtocolnum(), protocol2.getProtocolnum());
    
    List<Patient> patients1 = patientMapperMed1.selectByProtocolId(protocol1Id);
    List<Patient> patients2 = patientMapperMed2.selectByProtocolId(protocol2Id);
    
    List<PatienttoMRNBean> patienttoMRNs = new ArrayList<PatienttoMRNBean>();
    
    List<String> subjectIds = new ArrayList<String>();
    
    //計畫1中的病患ID暫存
    PatienttoMRNBean patienttoMRN=null;
    for(Patient patient1 : patients1) {
      patienttoMRN = new PatienttoMRNBean();
      patienttoMRN.setMrn(patient1.getMrn());
      patienttoMRN.setPatient1Id(patient1.getId());
      patienttoMRN.setPatient2Id(null);
      patienttoMRN.setSameFirstName(null);
      patienttoMRN.setSameLastName(null);
      patienttoMRNs.add(patienttoMRN);
      subjectIds.add(patient1.getMrn());
    }
    
    //兩計畫中的病患比對
    for(Patient patient2 : patients2) {
      //計畫2中有與計畫1一致的subject ID，將病患ID暫存在同個Index下
      if(subjectIds.contains(patient2.getMrn())) {
        patienttoMRNs.get(subjectIds.indexOf(patient2.getMrn())).setPatient2Id(patient2.getId());
      } else {
        patienttoMRN = new PatienttoMRNBean(); 
        patienttoMRN.setMrn(patient2.getMrn());
        patienttoMRN.setPatient1Id(null);
        patienttoMRN.setPatient2Id(patient2.getId());
        patienttoMRN.setSameFirstName(null);
        patienttoMRN.setSameLastName(null);
        patienttoMRNs.add(patienttoMRN);
        subjectIds.add(patient2.getMrn());
      }
    }
    
    //清除暫存
    patients1.removeAll(patients1);
    patients2.removeAll(patients2);
    subjectIds.removeAll(subjectIds);
    
    Patient patient1;
    Patient patient2;
    
    //開始病患基本資料比對
    for(int patientIndex = 0 ; patientIndex < patienttoMRNs.size() ; patientIndex++) {
      //確認病患ID均存在，有null代表其中一個計劃無此病患
      if(patienttoMRNs.get(patientIndex).getPatient1Id() != null && patienttoMRNs.get(patientIndex).getPatient2Id() != null) {
        //從CSIS資料庫中取得病患資料
        PatientExample ptest1 = new PatientExample();
        ptest1.or().andPatientidEqualTo(patienttoMRNs.get(patientIndex).getPatient1Id()).andProtocolidEqualTo(protocol1Id);
        PatientExample ptest2 = new PatientExample();
        ptest2.or().andPatientidEqualTo(patienttoMRNs.get(patientIndex).getPatient2Id()).andProtocolidEqualTo(protocol2Id);
        patient1 = patientMapperMed1.noStudyEventByPrimaryKey(ptest1);
        patient2 = patientMapperMed2.noStudyEventByPrimaryKey(ptest2);
        //比對病患資料是否一致，一致設定對應參數為true，反之，設定為false
        if( ( patient1.getFirstName() == null && patient2.getFirstName() == null ) 
            || ( "".equals(patient1.getFirstName()) && "".equals(patient2.getFirstName()) ) ) {
          patienttoMRNs.get(patientIndex).setSameFirstName(true);
        } else {
          if( patient1.getFirstName() != null && patient2.getFirstName() != null  
              && !"".equals(patient1.getFirstName()) && !"".equals(patient2.getFirstName()) ) {
            if(patient1.getFirstName().equals(patient2.getFirstName())) {
              patienttoMRNs.get(patientIndex).setSameFirstName(true);
            } else {
              patienttoMRNs.get(patientIndex).setSameFirstName(false);
            }
          } else {
            patienttoMRNs.get(patientIndex).setSameFirstName(false);
          }
        }
        if( ( patient1.getLastName() == null && patient2.getLastName() == null ) 
            || ( "".equals(patient1.getLastName()) && "".equals(patient2.getLastName()) ) ) {
          patienttoMRNs.get(patientIndex).setSameLastName(true);
        } else {
          if( patient1.getLastName() != null && patient2.getLastName() != null  
              && !"".equals(patient1.getLastName()) && !"".equals(patient2.getLastName()) ) {
            if(patient1.getLastName().equals(patient2.getLastName())) {
              patienttoMRNs.get(patientIndex).setSameLastName(true);
            } else {
              patienttoMRNs.get(patientIndex).setSameLastName(false);
            }
          } else {
            patienttoMRNs.get(patientIndex).setSameLastName(false);
          }
        }
      }
    }

    //執行病患分頁資料暫存函式
    ExportAction.exportpatient(writer, protocol1Id, protocol2Id, protocol1.getProtocolnum(), 
        protocol2.getProtocolnum(), patienttoMRNs, patientMapperMed1, patientMapperMed2);
    
    patienttoMRNs.removeAll(patienttoMRNs);
    
    List<Same_MRN> sameMRNs = Ebean.find(Same_MRN.class).findList();
    List<String> allVisit = null;
    List<PatientVisit> diffVisitDatePatients = new ArrayList<PatientVisit>();
    
    //開始訪視日期的比對
    for(int index = 0 ; index < sameMRNs.size() ; index++) {
      PatientExample ptest1 = new PatientExample();
      ptest1.or().andPatientidEqualTo(sameMRNs.get(index).getPatient1Id()).andProtocolidEqualTo(protocol1Id);
      ptest1.setOrderByClause("STUDYEVENTDEF.EVENTORDER");
      PatientExample ptest2 = new PatientExample();
      ptest2.or().andPatientidEqualTo(sameMRNs.get(index).getPatient2Id()).andProtocolidEqualTo(protocol2Id);
      ptest2.setOrderByClause("STUDYEVENTDEF.EVENTORDER");
      patient1 = patientMapperMed1.hasStudyEventByPrimaryKey(ptest1);
      patient2 = patientMapperMed2.hasStudyEventByPrimaryKey(ptest2);
      //暫存所有訪視，匯出時使用
      if(allVisit == null) {
        allVisit = new ArrayList<String>();
        for(StudyEvent studyEvent : patient1.getStudyEvents()) {
          allVisit.add(studyEvent.getName());
        }
      }
      //相同EVENTORDER下，比對訪視日期是否一致
      for(int studyEventIndex = 0 ; studyEventIndex < patient1.getStudyEvents().size() ; studyEventIndex++) {
        //判斷計畫1中訪視有沒有設定日期
        if( patient1.getStudyEvents().get(studyEventIndex).getVisitDate() != null ) {
          //判斷計畫2放式日期是否一致，不一致則匯出，一致則將資料儲存於暫存資料庫中
          if(!patient1.getStudyEvents().get(studyEventIndex).getVisitDate().equals(patient2.getStudyEvents().get(studyEventIndex).getVisitDate())) {
            PatientVisit patientVisit = new PatientVisit();
            patientVisit.setMrn(sameMRNs.get(index).getMrn());
            patientVisit.setEventOrder(patient1.getStudyEvents().get(studyEventIndex).getEventOrder());
            patientVisit.setVisit1Id(patient1.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setVisit2Id(patient2.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setVisitName(patient1.getStudyEvents().get(studyEventIndex).getName());
            patientVisit.setVisitDate1(patient1.getStudyEvents().get(studyEventIndex).getVisitDate());
            patientVisit.setVisitDate2(patient2.getStudyEvents().get(studyEventIndex).getVisitDate());
            diffVisitDatePatients.add(patientVisit);
          } else {
            PatientVisit patientVisit = new PatientVisit();
            patientVisit.setVisit1Id(patient1.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setVisit2Id(patient2.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setStudy1Id(patient1.getStudyEvents().get(studyEventIndex).getId());
            patientVisit.setStudy2Id(patient2.getStudyEvents().get(studyEventIndex).getId());
            patientVisit.setVisitName(patient1.getStudyEvents().get(studyEventIndex).getName());
            patientVisit.setVisitDate1(patient1.getStudyEvents().get(studyEventIndex).getVisitDate());
            patientVisit.setVisitDate2(patient2.getStudyEvents().get(studyEventIndex).getVisitDate());
            patientVisit.setMrnId(sameMRNs.get(index).getId());
            patientVisit.setMrn(sameMRNs.get(index).getMrn());
            Ebean.save(patientVisit);
          }
        } else {
          //當計畫1中訪視無設定日期，計畫2有設定，將會因不一致而匯出，均無設定日期則不匯出
          if(patient2.getStudyEvents().get(studyEventIndex).getVisitDate() != null) {
            PatientVisit patientVisit = new PatientVisit();
            patientVisit.setMrn(sameMRNs.get(index).getMrn());
            patientVisit.setEventOrder(patient1.getStudyEvents().get(studyEventIndex).getEventOrder());
            patientVisit.setVisit1Id(patient1.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setVisit2Id(patient2.getStudyEvents().get(studyEventIndex).getVisitId());
            patientVisit.setVisitName(patient1.getStudyEvents().get(studyEventIndex).getName());
            patientVisit.setVisitDate1(patient1.getStudyEvents().get(studyEventIndex).getVisitDate());
            patientVisit.setVisitDate2(patient2.getStudyEvents().get(studyEventIndex).getVisitDate());
            diffVisitDatePatients.add(patientVisit);
          }
        }
      }
    }
    
    sameMRNs.removeAll(sameMRNs);
    //執行訪視日期分頁資料暫存函式
    ExportAction.exportVisit(writer, protocol1.getProtocolnum(), protocol2.getProtocolnum(), diffVisitDatePatients, allVisit);
    if(allVisit != null)
      allVisit.removeAll(allVisit);
    diffVisitDatePatients.removeAll(diffVisitDatePatients);
    
    List<PatientVisit> patientVisits = Ebean.find(PatientVisit.class).findList();
    Map<String, Map<Integer, DataRecordBean>> crfDataRecords1 = new TreeMap<String, Map<Integer, DataRecordBean>>();
    Map<String, Map<Integer, DataRecordBean>> crfDataRecords2 = new TreeMap<String, Map<Integer, DataRecordBean>>();
    //儲存不同訪視中所有表單，Integer是為匯出階段重要指標
    Map<String, Map<String,Integer>> allCRFinVisit = new TreeMap<String, Map<String,Integer>>();
    List<CrfDatarecordBean> exportCRFStauts = new ArrayList<CrfDatarecordBean>();
    
    //開始表單狀態判斷
    for(int index = 0 ; index < patientVisits.size() ; index++) {
      crfDataRecords1.clear();
      crfDataRecords2.clear();
      //取得各個訪視內的表單(dataRecord)
      VisitExample visitExample1 = new VisitExample();
      visitExample1.or().andIdEqualTo(patientVisits.get(index).getVisit1Id())
      .andProtocolidEqualTo(protocol1Id);
      Visit visit1 = visitMapperMed1.selectStudyPlanVisits(visitExample1);
      VisitExample visitExample2 = new VisitExample();
      visitExample2.or().andIdEqualTo(patientVisits.get(index).getVisit2Id())
      .andProtocolidEqualTo(protocol2Id);
      Visit visit2 = visitMapperMed2.selectStudyPlanVisits(visitExample2);
      
      VisitExample visitFormExample1 = new VisitExample();
      visitFormExample1.or().andIdEqualTo(patientVisits.get(index).getVisit1Id());
      Visit visitForm1 = visitMapperMed1.selectVisits(visitFormExample1);
      if(visitForm1 != null)
        visit1.setFilloutForms(visitForm1.getFilloutForms());
      
      VisitExample visitFormExample2 = new VisitExample();
      visitFormExample2.or().andIdEqualTo(patientVisits.get(index).getVisit2Id());
      Visit visitForm2 = visitMapperMed2.selectVisits(visitFormExample2);
      if(visitForm2 != null)
        visit2.setFilloutForms(visitForm2.getFilloutForms());
      //將訪視內有填寫的表單依照表單名稱放入Map參數中，此步驟兩個計畫分開執行
      //計畫1
      for(ViewCRFBean crf1 : visit1.getFilloutForms()) {
        Map<Integer, DataRecordBean> dataRecords1 = null;
        //表單版本大於1，表單名稱後會加入版本
        if(crf1.getFormVersion() == 1) {
          dataRecords1 = crfDataRecords1.get(crf1.getTitle());
        } else {
          dataRecords1 = crfDataRecords1.get(crf1.getTitle()+" - v."+crf1.getFormVersion());
        }
        if(dataRecords1 == null) {
          dataRecords1 = new TreeMap<Integer, DataRecordBean>();
          if(crf1.getDataRecordId() != null) {
            DataRecordBean dataRecord = new DataRecordBean();
            dataRecord.setDataRecordId(crf1.getDataRecordId());
            dataRecord.setFormVersionId(crf1.getFormVersionId());
            dataRecord.setEventSequence(crf1.getEventSequence());
            dataRecord.setStatus(crf1.getStatus());
            dataRecords1.put(crf1.getEventSequence().intValue(), dataRecord);
            if(crf1.getFormVersion() == 1) {
              crfDataRecords1.put(crf1.getTitle(), dataRecords1);
            } else {
              crfDataRecords1.put(crf1.getTitle()+" - v."+crf1.getFormVersion(), dataRecords1);
            }
          } else {
            if(crf1.getFormVersion() == 1) {
              crfDataRecords1.put(crf1.getTitle(), dataRecords1);
            } else {
              crfDataRecords1.put(crf1.getTitle()+" - v."+crf1.getFormVersion(), dataRecords1);
            }
          }
        } else {
          DataRecordBean dataRecord = new DataRecordBean();
          dataRecord.setDataRecordId(crf1.getDataRecordId());
          dataRecord.setFormVersionId(crf1.getFormVersionId());
          dataRecord.setEventSequence(crf1.getEventSequence());
          dataRecord.setStatus(crf1.getStatus());
          dataRecords1.put(crf1.getEventSequence().intValue(), dataRecord);
          if(crf1.getFormVersion() == 1) {
            crfDataRecords1.put(crf1.getTitle(), dataRecords1);
          } else {
            crfDataRecords1.put(crf1.getTitle()+" - v."+crf1.getFormVersion(), dataRecords1);
          }
        }
      }

      //計畫2
      for(ViewCRFBean crf2 : visit2.getFilloutForms()) {
        Map<Integer, DataRecordBean> dataRecords2= null;
        //表單版本大於1，表單名稱後會加入版本
        if(crf2.getFormVersion() == 1) {
          dataRecords2 = crfDataRecords2.get(crf2.getTitle());
        } else {
          dataRecords2 = crfDataRecords2.get(crf2.getTitle()+" - v."+crf2.getFormVersion());
        }
        if(dataRecords2 == null) {
          dataRecords2 = new TreeMap<Integer, DataRecordBean>();
          if(crf2.getDataRecordId() != null) {
            DataRecordBean dataRecord = new DataRecordBean();
            dataRecord.setDataRecordId(crf2.getDataRecordId());
            dataRecord.setFormVersionId(crf2.getFormVersionId());
            dataRecord.setEventSequence(crf2.getEventSequence());
            dataRecord.setStatus(crf2.getStatus());
            dataRecords2.put(crf2.getEventSequence().intValue(), dataRecord);
            if(crf2.getFormVersion() == 1) {
              crfDataRecords2.put(crf2.getTitle(), dataRecords2);
            } else {
              crfDataRecords2.put(crf2.getTitle()+" - v."+crf2.getFormVersion(), dataRecords2);
            }
          } else {
            if(crf2.getFormVersion() == 1) {
              crfDataRecords2.put(crf2.getTitle(), dataRecords2);
            } else {
              crfDataRecords2.put(crf2.getTitle()+" - v."+crf2.getFormVersion(), dataRecords2);
            }
          }
        } else {
          DataRecordBean dataRecord = new DataRecordBean();
          dataRecord.setDataRecordId(crf2.getDataRecordId());
          dataRecord.setFormVersionId(crf2.getFormVersionId());
          dataRecord.setEventSequence(crf2.getEventSequence());
          dataRecord.setStatus(crf2.getStatus());
          dataRecords2.put(crf2.getEventSequence().intValue(), dataRecord);
          if(crf2.getFormVersion() == 1) {
            crfDataRecords2.put(crf2.getTitle(), dataRecords2);
          } else {
            crfDataRecords2.put(crf2.getTitle()+" - v."+crf2.getFormVersion(), dataRecords2);
          }
        }
      }
      
      //比較Map參數中表單種類多寡，確保每個種類都會被判斷到
      if(crfDataRecords1.size() >= crfDataRecords2.size()) {
        //當訪視1表單種類多，以此Map進行迴圈
        for(Map.Entry<String, Map<Integer, DataRecordBean>> entry1 : crfDataRecords1.entrySet()) {
          //將此表單放到對應的Map中
          Map<String,Integer> CRFsinVisit = allCRFinVisit.get(patientVisits.get(index).getVisitName());
          if(CRFsinVisit == null) {
            CRFsinVisit = new TreeMap<String,Integer>();
            CRFsinVisit.put(entry1.getKey(),null);
            allCRFinVisit.put(patientVisits.get(index).getVisitName(), CRFsinVisit);
          } else {
            CRFsinVisit.put(entry1.getKey(),null);
          }
          Map<Integer, DataRecordBean> entry2 = crfDataRecords2.get(entry1.getKey());
          if(entry2 != null && !entry2.isEmpty()) {
            //找尋訪視1和訪視2對應的表單
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry1.getValue().entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry1.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts1(dataRecordBeanentry.getValue().getStatus());
              crf.setDataRecord1Id(dataRecordBeanentry.getValue().getDataRecordId());
              crf.setFormVersion1Id(dataRecordBeanentry.getValue().getFormVersionId());
              if(entry2.get(dataRecordBeanentry.getKey()) != null) {
                crf.setDataRecord2Id(entry2.get(dataRecordBeanentry.getKey()).getDataRecordId());
                crf.setFormVersion2Id(entry2.get(dataRecordBeanentry.getKey()).getFormVersionId());
                crf.setStauts2(entry2.get(dataRecordBeanentry.getKey()).getStatus());
                crf.setMatchVisitId(patientVisits.get(index).getId());
                Ebean.save(crf);
                //移除已被對應的表單，剩下的為多出的表單
                entry2.remove(dataRecordBeanentry.getKey());
              } else {
                crf.setDataRecord2Id(null);
                crf.setFormVersion2Id(null);
                crf.setStauts2(ExportAction.formNotExist);
              }
              //表單狀態不一致則儲存在匯出List中
              if(!crf.getStauts1().equals(crf.getStauts2())) {
                if(excel != null) {
                  getCRFPage(excel, crf);
                } else {
                  crf.setPage(null);
                }
                exportCRFStauts.add(crf);
              }
            }
            //將多出的表單進行匯出紀錄
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry2.entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry1.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts1(ExportAction.formNotExist);
              crf.setStauts2(dataRecordBeanentry.getValue().getStatus());
              if(excel != null) {
                getCRFPage(excel, crf);
              } else {
                crf.setPage(null);
              }
              exportCRFStauts.add(crf);
            }
          } else {
            //訪視2無表單，直接將訪視1的表單進行匯出紀錄
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry1.getValue().entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry1.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts1(dataRecordBeanentry.getValue().getStatus());
              crf.setStauts2(ExportAction.formNotExist);
              if(excel != null) {
                getCRFPage(excel, crf);
              } else {
                crf.setPage(null);
              }
              exportCRFStauts.add(crf);
            }
          }
        }
      } else {
        //當訪視2表單種類多，以此Map進行迴圈
        for(Map.Entry<String, Map<Integer, DataRecordBean>> entry2 : crfDataRecords2.entrySet()) {
          //將此表單放到對應的Map中
          Map<String,Integer> CRFsinVisit = allCRFinVisit.get(patientVisits.get(index).getVisitName());
          if(CRFsinVisit == null) {
            CRFsinVisit = new TreeMap<String,Integer>();
            CRFsinVisit.put(entry2.getKey(),null);
            allCRFinVisit.put(patientVisits.get(index).getVisitName(), CRFsinVisit);
          } else {
            CRFsinVisit.put(entry2.getKey(),null);
          }
          Map<Integer, DataRecordBean> entry1 = crfDataRecords1.get(entry2.getKey());
          if(entry1 != null && !entry1.isEmpty()) {
            //找尋訪視2和訪視1對應的表單
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry2.getValue().entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry2.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts2(dataRecordBeanentry.getValue().getStatus());
              crf.setDataRecord2Id(dataRecordBeanentry.getValue().getDataRecordId());
              crf.setFormVersion2Id(dataRecordBeanentry.getValue().getFormVersionId());
              if(entry1.get(dataRecordBeanentry.getKey()) != null) {
                crf.setDataRecord1Id(entry1.get(dataRecordBeanentry.getKey()).getDataRecordId());
                crf.setFormVersion1Id(entry1.get(dataRecordBeanentry.getKey()).getFormVersionId());
                crf.setStauts1(entry1.get(dataRecordBeanentry.getKey()).getStatus());
                crf.setMatchVisitId(patientVisits.get(index).getId());
                Ebean.save(crf);
                //移除已被對應的表單，剩下的為多出的表單
                entry1.remove(dataRecordBeanentry.getKey());
              } else {
                crf.setDataRecord1Id(null);
                crf.setFormVersion1Id(null);
                crf.setStauts1(ExportAction.formNotExist);
              }
              //表單狀態不一致則儲存在匯出List中
              if(!crf.getStauts1().equals(crf.getStauts2())) {
                if(excel != null) {
                  getCRFPage(excel, crf);
                } else {
                  crf.setPage(null);
                }
                exportCRFStauts.add(crf);
              }
            }
            //將多出的表單進行匯出紀錄
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry1.entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry2.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts1(dataRecordBeanentry.getValue().getStatus());
              crf.setStauts2(ExportAction.formNotExist);
              if(excel != null) {
                getCRFPage(excel, crf);
              } else {
                crf.setPage(null);
              }
              exportCRFStauts.add(crf);
            }
          } else {
            //訪視1無表單，直接將訪視1的表單進行匯出紀錄
            for(Map.Entry<Integer, DataRecordBean> dataRecordBeanentry : entry2.getValue().entrySet()) {
              CrfDatarecordBean crf = new CrfDatarecordBean();
              crf.setTitle(entry2.getKey());
              crf.setMrn(patientVisits.get(index).getMrn());
              crf.setVisitName(patientVisits.get(index).getVisitName());
              crf.setEventSequence(dataRecordBeanentry.getValue().getEventSequence());
              crf.setStauts1(ExportAction.formNotExist);
              crf.setStauts2(dataRecordBeanentry.getValue().getStatus());
              if(excel != null) {
                getCRFPage(excel, crf);
              } else {
                crf.setPage(null);
              }
              exportCRFStauts.add(crf);
            }
          }
        }
      }
    }
    
    patientVisits.removeAll(patientVisits);
    //執行表單狀態分頁資料暫存函式
    ExportAction.exportCRFStauts(writer, exportCRFStauts, allCRFinVisit, protocol1.getProtocolnum(), protocol2.getProtocolnum());
    allCRFinVisit.clear();
    exportCRFStauts.removeAll(exportCRFStauts);
    
    //匯出檔案
    try {
      DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd");
      LocalDateTime now = LocalDateTime.now();
      writer.process(folderName, "DDE Report_Basic_"+dtf.format(now)+".xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
    
    ExportAction.exportLogPage(writerDP, protocol1.getProtocolnum(), protocol2.getProtocolnum());
    
    List<CrfDatarecordBean> CDB = Ebean.find(CrfDatarecordBean.class).findList();
    Map<String, List<PatientDatapoint>> exportDatapointMap = new TreeMap<String, List<PatientDatapoint>>();
    List<PatientDatapoint> exportDatapoints = null;
    
    boolean hasPageColumn;
    //Data Point比對判斷
    for(int index = 0 ; index < CDB.size() ; index++) {
      hasPageColumn = false;
      List<Datapoint> datapoints1 = datapointMapperMed1.selectByKey(CDB.get(index).getFormVersion1Id());
      List<Datapoint> datapoints2 = datapointMapperMed2.selectByKey(CDB.get(index).getFormVersion2Id());
      if(!datapoints1.isEmpty() && !datapoints2.isEmpty()) {
        for(int datapointIndex = 0 ; datapointIndex < datapoints1.size() ; datapointIndex++) {
          DatapointDatarecordExample DPDRExample1 = new DatapointDatarecordExample();
          DPDRExample1.or().andDatarecordidEqualTo(CDB.get(index).getDataRecord1Id()).andDatapointidEqualTo(datapoints1.get(datapointIndex).getId());
          List<DatapointDatarecord> DPDRE1 = datapointDatarecordMapperMed1.selectByExample(DPDRExample1);
          DatapointDatarecordExample DPDRExample2 = new DatapointDatarecordExample();
          DPDRExample2.or().andDatarecordidEqualTo(CDB.get(index).getDataRecord2Id()).andDatapointidEqualTo(datapoints2.get(datapointIndex).getId());
          List<DatapointDatarecord> DPDRE2 = datapointDatarecordMapperMed2.selectByExample(DPDRExample2);
          if( !DPDRE1.isEmpty() && !DPDRE2.isEmpty()) {
            if(!DPDRE1.get(0).getValue().equals(DPDRE2.get(0).getValue())) {
              PatientDatapoint exportDatapoint = new PatientDatapoint();
              String DP_Page = getDPPage(excel, CDB.get(index), datapoints1.get(datapointIndex).getName());
              exportDatapoint.setPage(DP_Page);
              if(DP_Page != null) {
                hasPageColumn = true;
              }
              exportDatapoint.setMrn(CDB.get(index).getMrn());
              if(CDB.get(index).getEventSequence() != 1) {
                exportDatapoint.setFormTitle(CDB.get(index).getTitle()+" | #"+CDB.get(index).getEventSequence());
              } else {
                exportDatapoint.setFormTitle(CDB.get(index).getTitle());
              }
              exportDatapoint.setVisitName(CDB.get(index).getVisitName());
              exportDatapoint.setDatapointName(datapoints1.get(datapointIndex).getName());
              exportDatapoint.setQuestionText(datapoints1.get(datapointIndex).getText());
              if(datapoints1.get(datapointIndex).getEnumType().isEmpty() && datapoints2.get(datapointIndex).getEnumType().isEmpty()) {
                exportDatapoint.setValue1(DPDRE1.get(0).getValue());
                exportDatapoint.setValue2(DPDRE2.get(0).getValue());
              } else {
                for(EnumType enumType : datapoints1.get(datapointIndex).getEnumType()) {
                  if(DPDRE1.get(0).getValue().equals(enumType.getValue())) {
                    exportDatapoint.setValue1(enumType.getDescription());
                    break;
                  }
                }
                for(EnumType enumType : datapoints2.get(datapointIndex).getEnumType()) {
                  if(DPDRE2.get(0).getValue().equals(enumType.getValue())) {
                    exportDatapoint.setValue2(enumType.getDescription());
                    break;
                  }
                }
              }
              exportDatapoints = exportDatapointMap.get(CDB.get(index).getVisitName());
              if(exportDatapoints != null) {
                exportDatapoints.add(exportDatapoint);
                if(hasPageColumn && hasPageColumnMap.get(CDB.get(index).getVisitName()) == null) {
                  hasPageColumnMap.put(CDB.get(index).getVisitName(), hasPageColumn);
                }
              } else {
                exportDatapoints = new ArrayList<PatientDatapoint>();
                exportDatapoints.add(exportDatapoint);
                exportDatapointMap.put(CDB.get(index).getVisitName(), exportDatapoints);
                hasPageColumnMap.put(CDB.get(index).getVisitName(), hasPageColumn);
              }
            }
          } else {
            if( (!DPDRE1.isEmpty() && DPDRE2.isEmpty()) || (DPDRE1.isEmpty() && !DPDRE2.isEmpty()) ) {
              if(!DPDRE1.isEmpty()) {
                PatientDatapoint exportDatapoint = new PatientDatapoint();
                String DP_Page = getDPPage(excel, CDB.get(index), datapoints1.get(datapointIndex).getName());
                exportDatapoint.setPage(DP_Page);
                if(DP_Page != null) {
                  hasPageColumn = true;
                }
                exportDatapoint.setMrn(CDB.get(index).getMrn());
                if(CDB.get(index).getEventSequence() != 1) {
                  exportDatapoint.setFormTitle(CDB.get(index).getTitle()+" | #"+CDB.get(index).getEventSequence());
                } else {
                  exportDatapoint.setFormTitle(CDB.get(index).getTitle());
                }
                exportDatapoint.setVisitName(CDB.get(index).getVisitName());
                exportDatapoint.setDatapointName(datapoints1.get(datapointIndex).getName());
                exportDatapoint.setQuestionText(datapoints1.get(datapointIndex).getText());
                if( datapoints1.get(datapointIndex).getEnumType().isEmpty() ) {
                  exportDatapoint.setValue1(DPDRE1.get(0).getValue());
                } else {
                  for(EnumType enumType : datapoints1.get(datapointIndex).getEnumType()) {
                    if(DPDRE1.get(0).getValue().equals(enumType.getValue())) {
                      exportDatapoint.setValue1(enumType.getDescription());
                      break;
                    }
                  }
                }
                exportDatapoint.setValue2(null);
                exportDatapoints = exportDatapointMap.get(CDB.get(index).getVisitName());
                if(exportDatapoints != null) {
                  exportDatapoints.add(exportDatapoint);
                  if(hasPageColumn && hasPageColumnMap.get(CDB.get(index).getVisitName()) == null) {
                    hasPageColumnMap.put(CDB.get(index).getVisitName(), hasPageColumn);
                  }
                } else {
                  exportDatapoints = new ArrayList<PatientDatapoint>();
                  exportDatapoints.add(exportDatapoint);
                  exportDatapointMap.put(CDB.get(index).getVisitName(), exportDatapoints);
                  hasPageColumnMap.put(CDB.get(index).getVisitName(), hasPageColumn);
                }
              } else {
                PatientDatapoint exportDatapoint = new PatientDatapoint();
                String DP_Page = getDPPage(excel, CDB.get(index), datapoints1.get(datapointIndex).getName());
                exportDatapoint.setPage(DP_Page);
                if(DP_Page != null) {
                  hasPageColumn = true;
                }
                exportDatapoint.setMrn(CDB.get(index).getMrn());
                if(CDB.get(index).getEventSequence() != 1) {
                  exportDatapoint.setFormTitle(CDB.get(index).getTitle()+" | #"+CDB.get(index).getEventSequence());
                } else {
                  exportDatapoint.setFormTitle(CDB.get(index).getTitle());
                }
                exportDatapoint.setVisitName(CDB.get(index).getVisitName());
                exportDatapoint.setDatapointName(datapoints1.get(datapointIndex).getName());
                exportDatapoint.setQuestionText(datapoints1.get(datapointIndex).getText());
                exportDatapoint.setValue1(null);
                if( datapoints2.get(datapointIndex).getEnumType().isEmpty() ) {
                  exportDatapoint.setValue2(DPDRE2.get(0).getValue());
                } else {
                  for(EnumType enumType : datapoints2.get(datapointIndex).getEnumType()) {
                    if(DPDRE2.get(0).getValue().equals(enumType.getValue())) {
                      exportDatapoint.setValue2(enumType.getDescription());
                      break;
                    }
                  }
                }
                exportDatapoints = exportDatapointMap.get(CDB.get(index).getVisitName());
                if(exportDatapoints != null) {
                  exportDatapoints.add(exportDatapoint);
                  if(hasPageColumn && hasPageColumnMap.get(CDB.get(index).getVisitName()) == null ){
                    hasPageColumnMap.put(CDB.get(index).getVisitName(), hasPageColumn);
                  }
                } else {
                  exportDatapoints = new ArrayList<PatientDatapoint>();
                  exportDatapoints.add(exportDatapoint);
                  exportDatapointMap.put(CDB.get(index).getVisitName(), exportDatapoints);
                }
              }
            }
          }
        }
      }
    }
    
    CDB.removeAll(CDB);
    if( !exportDatapointMap.isEmpty() ) {
      for(Map.Entry<String, List<PatientDatapoint>> entry : exportDatapointMap.entrySet()) {
        ExportAction.exportDatapoint(writerDP, entry.getValue(), protocol1.getProtocolnum(), protocol2.getProtocolnum(), entry.getKey(), hasPageColumnMap.get(entry.getKey()));
      }
    }
    exportDatapointMap.clear();
    
    //匯出檔案
    try {
      DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyyMMdd");
      LocalDateTime now = LocalDateTime.now();
      writerDP.process(folderName, "DDE Report_Form_"+dtf.format(now)+".xlsx");
    } catch (Exception e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
    
    //刪除暫存資料庫中所有資料
    Ebean.delete(Ebean.find(Same_MRN.class).findList());
    Ebean.delete(Ebean.find(PatientVisit.class).findList());
    Ebean.delete(Ebean.find(CrfDatarecordBean.class).findList());
    
    mybatisConnect.endSqlSession();
  }
  
  public String getDPPage(List<List<List<String>>> excel, CrfDatarecordBean CDB, String DP_Name) {
    String DP_Page = null;
//    String form_Name = null;
    
    if(excel == null) {
      return DP_Page;
    }
    
    int rowCounter;
//    if(CDB.getEventSequence() == 1) {
//      form_Name = CDB.getTitle();
//    } else {
//      form_Name = CDB.getTitle()+" | #"+CDB.getEventSequence();
//    }
    for(List<List<String>> sheet : excel) {
      if(sheet.get(1).get(1).equals(CDB.getTitle())) {
        rowCounter = 0;
        for(List<String> row : sheet) {
          if( rowCounter == 0 ) {
            rowCounter++;
            continue;
          }
          if(row.get(0).equals(CDB.getVisitName()) && row.get(2).equals(DP_Name)) {
            DP_Page = row.get(3);
            break;
          }
        }
        break;
      }
    }
    return DP_Page;
  }
  
  public void getCRFPage(List<List<List<String>>> excel, CrfDatarecordBean crf) {
    String form_Name = null;
    int rowCounter;
    if(crf.getEventSequence() == 1) {
      form_Name = crf.getTitle();
    } else {
      form_Name = crf.getTitle()+" | #"+crf.getEventSequence();
    }
    for(List<List<String>> sheet : excel) {
      if(sheet.get(1).get(1).equals(form_Name)) {
        rowCounter = 0;
        for(List<String> row : sheet) {
          if( rowCounter == 0 ) {
            rowCounter++;
            continue;
          }
          if(row.get(0).equals(crf.getVisitName())) {
            crf.setPage(row.get(3));
            break;
          }
        }
        break;
      }
    }
  }
  
}
