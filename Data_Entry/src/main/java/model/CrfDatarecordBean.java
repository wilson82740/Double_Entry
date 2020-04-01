package model;

public class CrfDatarecordBean {
  
  private long id;

  private String title;
  
  private Long dataRecord1Id;
 
  private Long formVersion1Id;
  
  private Long dataRecord2Id;
  
  private Long formVersion2Id;

  private long matchVisitId;

  private String visitName;

  private String mrn;
  
  private Long eventSequence;
  
  private String Stauts1;
  
  private String Stauts2;
  
  public long getId() {
    return id;
  }

  public void setId(long id) {
    this.id = id;
  }

  public String getTitle() {
    return title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public Long getDataRecord1Id() {
    return dataRecord1Id;
  }

  public void setDataRecord1Id(Long dataRecord1Id) {
    this.dataRecord1Id = dataRecord1Id;
  }

  public Long getFormVersion1Id() {
    return formVersion1Id;
  }

  public void setFormVersion1Id(Long formVersion1Id) {
    this.formVersion1Id = formVersion1Id;
  }

  public Long getDataRecord2Id() {
    return dataRecord2Id;
  }

  public void setDataRecord2Id(Long dataRecord2Id) {
    this.dataRecord2Id = dataRecord2Id;
  }

  public Long getFormVersion2Id() {
    return formVersion2Id;
  }

  public void setFormVersion2Id(Long formVersion2Id) {
    this.formVersion2Id = formVersion2Id;
  }

  public long getMatchVisitId() {
    return matchVisitId;
  }

  public void setMatchVisitId(long matchVisitId) {
    this.matchVisitId = matchVisitId;
  }

  public String getVisitName() {
    return visitName;
  }

  public void setVisitName(String visitName) {
    this.visitName = visitName;
  }

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public Long getEventSequence() {
    return eventSequence;
  }

  public void setEventSequence(Long eventSequence) {
    this.eventSequence = eventSequence;
  }

  public String getStauts1() {
    return Stauts1;
  }

  public void setStauts1(String stauts1) {
    Stauts1 = stauts1;
  }

  public String getStauts2() {
    return Stauts2;
  }

  public void setStauts2(String stauts2) {
    Stauts2 = stauts2;
  }
  
}
