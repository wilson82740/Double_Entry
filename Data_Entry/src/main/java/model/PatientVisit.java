package model;

import java.util.Date;
import java.util.List;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Table(name = "MATCH_VISIT")
public class PatientVisit {
  
  @Id
  private Long id;
  
  @Column(name = "MRNID")
  private Long mrnId;
  
  private String mrn;

  private Long study1Id;
  
  private Long study2Id;
  
  @Column(name = "EVENTORDER")
  private Long eventOrder;

  private Long visit1Id;
  
  private Long visit2Id;

  @Column(name = "VISITNAME")
  private String visitName;

  @Column(name = "VISITDATE1")
  private Date visitDate1;
  
  @Column(name = "VISITDATE2")
  private Date visitDate2;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public Long getMrnId() {
    return mrnId;
  }

  public void setMrnId(Long mrnId) {
    this.mrnId = mrnId;
  }

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public Long getStudy1Id() {
    return study1Id;
  }

  public void setStudy1Id(Long study1Id) {
    this.study1Id = study1Id;
  }

  public Long getStudy2Id() {
    return study2Id;
  }

  public void setStudy2Id(Long study2Id) {
    this.study2Id = study2Id;
  }

  public Long getEventOrder() {
    return eventOrder;
  }

  public void setEventOrder(Long eventOrder) {
    this.eventOrder = eventOrder;
  }

  public Long getVisit1Id() {
    return visit1Id;
  }

  public void setVisit1Id(Long visit1Id) {
    this.visit1Id = visit1Id;
  }

  public Long getVisit2Id() {
    return visit2Id;
  }

  public void setVisit2Id(Long visit2Id) {
    this.visit2Id = visit2Id;
  }

  public String getVisitName() {
    return visitName;
  }

  public void setVisitName(String visitName) {
    this.visitName = visitName;
  }

  public Date getVisitDate1() {
    return visitDate1;
  }

  public void setVisitDate1(Date visitDate1) {
    this.visitDate1 = visitDate1;
  }

  public Date getVisitDate2() {
    return visitDate2;
  }

  public void setVisitDate2(Date visitDate2) {
    this.visitDate2 = visitDate2;
  }

}
