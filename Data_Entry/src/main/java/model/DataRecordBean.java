package model;

public class DataRecordBean {

  private long dataRecordId;

  private long formVersionId;

  private long eventSequence;
  
  private String status;

  public long getDataRecordId() {
    return dataRecordId;
  }

  public void setDataRecordId(long dataRecordId) {
    this.dataRecordId = dataRecordId;
  }

  public long getFormVersionId() {
    return formVersionId;
  }

  public void setFormVersionId(long formVersionId) {
    this.formVersionId = formVersionId;
  }

  public long getEventSequence() {
    return eventSequence;
  }

  public void setEventSequence(long eventSequence) {
    this.eventSequence = eventSequence;
  }

  public String getStatus() {
    return status;
  }

  public void setStatus(String status) {
    this.status = status;
  }

}
