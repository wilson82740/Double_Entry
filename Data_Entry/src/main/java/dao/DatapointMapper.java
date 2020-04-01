package dao;

import java.util.List;

import model.Datapoint;

public interface DatapointMapper {
  List<Datapoint> selectByKey(Long id);
}
