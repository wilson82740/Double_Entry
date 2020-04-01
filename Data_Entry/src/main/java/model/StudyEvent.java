package model;

import java.util.Date;

public class StudyEvent {

  private Long id;

  private String name;

  private Date visitDate;
  
  private Long visitId;
  
  private Long eventOrder;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public String getName() {
    return name;
  }

  public void setName(String name) {
    this.name = name;
  }

  public Date getVisitDate() {
    return visitDate;
  }

  public void setVisitDate(Date visitDate) {
    this.visitDate = visitDate;
  }

  public Long getVisitId() {
    return visitId;
  }

  public void setVisitId(Long visitId) {
    this.visitId = visitId;
  }

  public Long getEventOrder() {
    return eventOrder;
  }

  public void setEventOrder(Long eventOrder) {
    this.eventOrder = eventOrder;
  }

}
