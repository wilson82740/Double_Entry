package model;

public class PatienttoMRNBean {
  
  private String mrn;
  
  private Long patient1Id;

  private Long patient2Id;
  
  private Boolean sameFirstName;
  
  private Boolean sameLastName;

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public Long getPatient1Id() {
    return patient1Id;
  }

  public void setPatient1Id(Long patient1Id) {
    this.patient1Id = patient1Id;
  }

  public Long getPatient2Id() {
    return patient2Id;
  }

  public void setPatient2Id(Long patient2Id) {
    this.patient2Id = patient2Id;
  }

  public Boolean getSameFirstName() {
    return sameFirstName;
  }

  public void setSameFirstName(Boolean sameFirstName) {
    this.sameFirstName = sameFirstName;
  }

  public Boolean getSameLastName() {
    return sameLastName;
  }

  public void setSameLastName(Boolean sameLastName) {
    this.sameLastName = sameLastName;
  }
  
}
