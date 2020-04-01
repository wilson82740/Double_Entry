package dao;

import java.util.List;

import model.Same_MRN;

public interface Same_MRNMapper {
  
  List<Same_MRN> selectAll();
  
  int insert(Same_MRN sameMRN);
  
  int deleteAll();
}
