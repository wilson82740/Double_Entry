package model;

import java.util.List;

public class FormVersionTest {
  
  private Long id;
  
  private Long formVersion;
  
  private List<SectionTest> sectionTest;
  
  public Long getId() {
    return id;
  }

  public void setId(Long id) {
    this.id = id;
  }

  public Long getFormVersion() {
    return formVersion;
  }

  public void setFormVersion(Long formVersion) {
    this.formVersion = formVersion;
  }

  public List<SectionTest> getSectionTest() {
    return sectionTest;
  }

  public void setSectionTest(List<SectionTest> sectionTest) {
    this.sectionTest = sectionTest;
  }
  
}
