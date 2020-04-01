package dao;

import model.DatapointDatarecord;
import model.DatapointDatarecordExample;
import java.util.List;
import org.apache.ibatis.annotations.Param;

public interface DatapointDatarecordMapper {
    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int countByExample(DatapointDatarecordExample example);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int deleteByExample(DatapointDatarecordExample example);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int insert(DatapointDatarecord record);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int insertSelective(DatapointDatarecord record);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    List<DatapointDatarecord> selectByExample(DatapointDatarecordExample example);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int updateByExampleSelective(@Param("record") DatapointDatarecord record, @Param("example") DatapointDatarecordExample example);

    /**
     * This method was generated by MyBatis Generator.
     * This method corresponds to the database table DATAPOINTDATARECORD
     *
     * @mbggenerated
     */
    int updateByExample(@Param("record") DatapointDatarecord record, @Param("example") DatapointDatarecordExample example);
}