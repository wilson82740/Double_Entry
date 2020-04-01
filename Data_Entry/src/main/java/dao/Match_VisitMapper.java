package dao;

import java.util.List;

import model.PatientVisit;

public interface Match_VisitMapper {
  
  List<PatientVisit> selectAll();

  int insert(PatientVisit patientVisit);
  
  int deleteAll();
}
