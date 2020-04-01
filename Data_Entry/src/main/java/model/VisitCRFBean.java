package model;

import java.util.List;

public class VisitCRFBean {
  
  private long visitId;
  
  private String name;
  
  private List<CRFBean> crfs;

  public long getVisitId() {
    return visitId;
  }

  public void setVisitId(long visitId) {
    this.visitId = visitId;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public List<CRFBean> getCrfs() {
    return crfs;
  }

  public void setCrfs(List<CRFBean> crfs) {
    this.crfs = crfs;
  }
  
}
