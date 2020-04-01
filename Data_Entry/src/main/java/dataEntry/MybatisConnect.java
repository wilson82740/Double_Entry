package dataEntry;

import java.io.IOException;
import java.io.InputStream;

import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

public class MybatisConnect {

  protected SqlSession sqlSessionCsis;
  protected SqlSession sqlSessionDataEntry;
  
  public SqlSession getSqlSessionCsis() {
    return sqlSessionCsis;
  }

  public SqlSession getSqlSessionDataEntry() {
    return sqlSessionDataEntry;
  }

  public void setSqlSession(String resourceCsis, String resourceDataEntry) throws IOException {
    InputStream inputStreamCsis = Resources.getResourceAsStream(resourceCsis);

    SqlSessionFactory sqlSessionFactoryCsis =
        new SqlSessionFactoryBuilder().build(inputStreamCsis);
    
    sqlSessionCsis = sqlSessionFactoryCsis.openSession();
    
    InputStream inputStreamDataEntry = Resources.getResourceAsStream(resourceDataEntry);

    SqlSessionFactory sqlSessionFactoryDataEntry =
        new SqlSessionFactoryBuilder().build(inputStreamDataEntry);
    
    sqlSessionDataEntry = sqlSessionFactoryDataEntry.openSession();
    
  }
  
  public void endSqlSession() {
    sqlSessionCsis.commit();
    sqlSessionCsis.close();
    sqlSessionDataEntry.commit();
    sqlSessionDataEntry.close();
  }
  
}
