package model;

public class ViewCRFBean {
private Long id;
  
  private Long dataRecordId;
  
  private Long formVersionId;
  
  private Long formVersion;
  
  private String title;
  
  private Long eventSequence;
  
  private String status;

  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public Long getDataRecordId() {
    return dataRecordId;
  }

  public void setDataRecordId(Long dataRecordId) {
    this.dataRecordId = dataRecordId;
  }

  public Long getFormVersionId() {
    return formVersionId;
  }

  public void setFormVersionId(Long formVersionId) {
    this.formVersionId = formVersionId;
  }

  public Long getFormVersion() {
    return formVersion;
  }

  public void setFormVersion(Long formVersion) {
    this.formVersion = formVersion;
  }

  public String getTitle() {
    return title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public Long getEventSequence() {
    return eventSequence;
  }

  public void setEventSequence(Long eventSequence) {
    this.eventSequence = eventSequence;
  }

  public String getStatus() {
    return status;
  }

  public void setStatus(String status) {
    this.status = status;
  }
  
}
