package dao;

import java.util.List;

import model.Visit;
import model.VisitExample;

public interface VisitMapper {
  Visit selectStudyPlanVisits(VisitExample example);
  
  Visit selectVisits(VisitExample example);
}
