package model;

public class Same_MRN {

  private Long id;
  
  private String mrn;

  private Long patient1Id;

  private Long patient2Id;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

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

}
