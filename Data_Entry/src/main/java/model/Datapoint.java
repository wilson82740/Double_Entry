package model;

import java.util.List;

public class Datapoint {

  private Long id;
  
  private String name;
  
  private String text;
  
  private Long measurementUnitId;
  
  private String displayType;
  
  private List<EnumType> enumType;

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

  public String getText() {
    return text;
  }

  public void setText(String text) {
    this.text = text;
  }

  public Long getMeasurementUnitId() {
    return measurementUnitId;
  }

  public void setMeasurementUnitId(Long measurementUnitId) {
    this.measurementUnitId = measurementUnitId;
  }

  public String getDisplayType() {
    return displayType;
  }

  public void setDisplayType(String displayType) {
    this.displayType = displayType;
  }

  public List<EnumType> getEnumType() {
    return enumType;
  }

  public void setEnumType(List<EnumType> enumType) {
    this.enumType = enumType;
  }

}
