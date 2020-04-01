package dao;

import java.util.List;

import model.Patient;
import model.PatientExample;

public interface PatientMapper {
  List<Patient> selectByProtocolId(Long protocolid);
  
  Patient noStudyEventByPrimaryKey(PatientExample example);
  
  Patient hasStudyEventByPrimaryKey(PatientExample example);
}
