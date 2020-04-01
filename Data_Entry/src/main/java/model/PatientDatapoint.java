package model;

public class PatientDatapoint {

  private String formTitle;
  
  private String visitName;

  private String mrn;
  
  private String questionText;
  
  private String value1;
  
  private String value2;

  public String getFormTitle() {
    return formTitle;
  }

  public void setFormTitle(String formTitle) {
    this.formTitle = formTitle;
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

  public String getQuestionText() {
    return questionText;
  }

  public void setQuestionText(String questionText) {
    this.questionText = questionText;
  }

  public String getValue1() {
    return value1;
  }

  public void setValue1(String value1) {
    this.value1 = value1;
  }

  public String getValue2() {
    return value2;
  }

  public void setValue2(String value2) {
    this.value2 = value2;
  }
  
}
