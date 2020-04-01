package model;

import java.util.List;

public class Visit {
  private Long id;
  
  private Long patientId;
  
  private String mrn;
  
  private List<ViewCRFBean> filloutForms;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public Long getPatientId() {
    return patientId;
  }

  public void setPatientId(Long patientId) {
    this.patientId = patientId;
  }

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public List<ViewCRFBean> getFilloutForms() {
    return filloutForms;
  }

  public void setFilloutForms(List<ViewCRFBean> filloutForms) {
    this.filloutForms = filloutForms;
  }
  
}
