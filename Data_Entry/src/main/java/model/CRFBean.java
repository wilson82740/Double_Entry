package model;

import java.util.List;

public class CRFBean {
  
  private long id;
  
  private String title;
  
  private int number;

  private List<DataRecordBean> DataRecords;
  
  private List<Datapoint> DataPoints;
  
  public long getId() {
    return id;
  }

  public void setId(long id) {
    this.id = id;
  }

  public String getTitle() {
    return title;
  }

  public void setTitle(String title) {
    this.title = title;
  }

  public int getNumber() {
    return number;
  }

  public void setNumber(int number) {
    this.number = number;
  }

  public List<DataRecordBean> getDataRecords() {
    return DataRecords;
  }

  public void setDataRecords(List<DataRecordBean> dataRecords) {
    DataRecords = dataRecords;
  }

  public List<Datapoint> getDataPoints() {
    return DataPoints;
  }

  public void setDataPoints(List<Datapoint> dataPoints) {
    DataPoints = dataPoints;
  }
  
}
