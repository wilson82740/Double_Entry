package model;

import java.util.List;

public class Patient {
  private Long id;
  
  private Long protocolid;
  
  private String mrn;
  
  private String lastName;
  
  private String firstName;
  
  private List<StudyEvent> studyEvents;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public Long getProtocolid() {
    return protocolid;
  }

  public void setProtocolid(Long protocolid) {
    this.protocolid = protocolid;
  }

  public String getMrn() {
    return mrn;
  }

  public void setMrn(String mrn) {
    this.mrn = mrn;
  }

  public String getLastName() {
    return lastName;
  }

  public void setLastName(String lastName) {
    this.lastName = lastName;
  }

  public String getFirstName() {
    return firstName;
  }

  public void setFirstName(String firstName) {
    this.firstName = firstName;
  }

  public List<StudyEvent> getStudyEvents() {
    return studyEvents;
  }

  public void setStudyEvents(List<StudyEvent> studyEvents) {
    this.studyEvents = studyEvents;
  }
  
}
