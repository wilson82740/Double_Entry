package model;

public class PatientCRF {

  private String Title;

  private int number1;
  
  private int number2;

  private long dataRecord1Id;
  
  private long dataRecord2Id;

  private long formVersion1Id;
  
  private long formVersion2Id;

  private String visitName;

  private String mrn;

  public String getTitle() {
    return Title;
  }

  public void setTitle(String title) {
    Title = title;
  }

  public int getNumber1() {
    return number1;
  }

  public void setNumber1(int number1) {
    this.number1 = number1;
  }

  public int getNumber2() {
    return number2;
  }

  public void setNumber2(int number2) {
    this.number2 = number2;
  }

  public long getDataRecord1Id() {
    return dataRecord1Id;
  }

  public void setDataRecord1Id(long dataRecord1Id) {
    this.dataRecord1Id = dataRecord1Id;
  }

  public long getDataRecord2Id() {
    return dataRecord2Id;
  }

  public void setDataRecord2Id(long dataRecord2Id) {
    this.dataRecord2Id = dataRecord2Id;
  }

  public long getFormVersion1Id() {
    return formVersion1Id;
  }

  public void setFormVersion1Id(long formVersion1Id) {
    this.formVersion1Id = formVersion1Id;
  }

  public long getFormVersion2Id() {
    return formVersion2Id;
  }

  public void setFormVersion2Id(long formVersion2Id) {
    this.formVersion2Id = formVersion2Id;
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

}
