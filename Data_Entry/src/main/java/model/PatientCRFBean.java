package model;

import java.util.List;

public class PatientCRFBean {
  private long protocolId;

  private long patientId;

  private String mrn;

  private List<VisitCRFBean> visits;

  public long getProtocolId() {
    return protocolId;
  }

  public void setProtocolId(long protocolId) {
    this.protocolId = protocolId;
  }

  public long getPatientId() {
    return patientId;
  }

  public void setPatientId(long patientId) {
    this.patientId = patientId;
  }

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public List<VisitCRFBean> getVisits() {
    return visits;
  }

  public void setVisits(List<VisitCRFBean> visits) {
    this.visits = visits;
  }

}
