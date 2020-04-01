package dao;

import model.Form;
import model.FormExample;

import java.util.List;
import org.apache.ibatis.annotations.Param;

public interface FormMapper {
  Form selectByPrimaryKey(Long id);
  Form selectByPrimaryKey1(Long id);
}