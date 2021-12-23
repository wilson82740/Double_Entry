package dataEntry;

import java.io.IOException;
import java.io.InputStream;

import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

import com.avaje.ebean.EbeanServerFactory;
import com.avaje.ebean.config.DataSourceConfig;
import com.avaje.ebean.config.ServerConfig;

import model.CrfDatarecordBean;
import model.PatientVisit;
import model.Same_MRN;

public class MybatisConnect {

  protected SqlSession sqlSessionMed1;
  protected SqlSession sqlSessionMed2;
  protected SqlSession sqlSessionDataEntry;
  
  public SqlSession getSqlSessionMed1() {
    return sqlSessionMed1;
  }
  
  public SqlSession getSqlSessionMed2() {
    return sqlSessionMed2;
  }

  public SqlSession getSqlSessionDataEntry() {
    return sqlSessionDataEntry;
  }

  public void setSqlSession(String resourceMed1, String resourceMed2, String resourceDataEntry) throws IOException {
    InputStream inputStreamMed1 = Resources.getResourceAsStream(resourceMed1);

    SqlSessionFactory sqlSessionFactoryMed1 =
        new SqlSessionFactoryBuilder().build(inputStreamMed1);
    
    sqlSessionMed1 = sqlSessionFactoryMed1.openSession();
    
    InputStream inputStreamMed2 = Resources.getResourceAsStream(resourceMed2);

    SqlSessionFactory sqlSessionFactoryMed2 =
        new SqlSessionFactoryBuilder().build(inputStreamMed2);
    
    sqlSessionMed2 = sqlSessionFactoryMed2.openSession();
    
//    InputStream inputStreamDataEntry = Resources.getResourceAsStream(resourceDataEntry);
//
//    SqlSessionFactory sqlSessionFactoryDataEntry =
//        new SqlSessionFactoryBuilder().build(inputStreamDataEntry);
//    
//    sqlSessionDataEntry = sqlSessionFactoryDataEntry.openSession();
    ServerConfig config = new ServerConfig();
    config.setName("double_entry");

    DataSourceConfig h2Db = new DataSourceConfig();
    h2Db.setDriver("org.h2.Driver");
    h2Db.setUsername("double_entry");
    h2Db.setPassword("0000");
    h2Db.setUrl("jdbc:h2:double_entry");
    h2Db.setHeartbeatSql("select 1");

    config.setDataSourceConfig(h2Db);

    config.setDdlGenerate(false);
    config.setDdlRun(false);

    config.setDefaultServer(true);
    config.setRegister(true);

    config.addClass(Same_MRN.class);
    config.addClass(PatientVisit.class);
    config.addClass(CrfDatarecordBean.class);

    EbeanServerFactory.create(config);
    
  }
  
  public void endSqlSession() {
    sqlSessionMed1.commit();
    sqlSessionMed1.close();
    sqlSessionMed2.commit();
    sqlSessionMed2.close();
//    sqlSessionDataEntry.commit();
//    sqlSessionDataEntry.close();
  }
  
}
