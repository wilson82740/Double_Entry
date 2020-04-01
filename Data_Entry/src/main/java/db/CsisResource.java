package db;

import java.io.IOException;
import java.io.Reader;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

/**
 * 
 * @author Wei-Ming Wu
 * 
 */
public enum CsisResource {

  CSIS("CsisMapperConfig.xml", "production");

  private final String resource;
  private final String environment;
  private final SqlSessionFactory sessionFactory;

  private CsisResource(String resource, String environment) {
    this.resource = resource;
    this.environment = environment;
    Reader resourceReader = getResource();
    sessionFactory = new SqlSessionFactoryBuilder().build(resourceReader);
    try {
      resourceReader.close();
    } catch (IOException e) {
      Logger.getLogger(this.getClass().getName()).log(Level.SEVERE, null, e);
    }
  }

  public Reader getResource() {
    Reader reader = null;

    try {
      reader = Resources.getResourceAsReader(resource);
    } catch (IOException ex) {
      Logger.getLogger(CsisResource.class.getName())
          .log(Level.SEVERE, null, ex);
    }

    return reader;
  }

  public String getEnvironment() {
    return environment;
  }

  public SqlSessionFactory getSessionFactory() {
    return sessionFactory;
  }

}
