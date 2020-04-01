package dao;

import java.util.List;

import model.CrfDatarecordBean;

public interface Crf_DatarecordMapper {

  List<CrfDatarecordBean> selectAll();

  int insert(CrfDatarecordBean CDBean);
  
  int deleteAll();
}
