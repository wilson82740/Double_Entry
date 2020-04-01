package db;

import org.apache.ibatis.session.SqlSessionFactory;

public final class CsisConnectionManager {

  private static volatile CsisConnectionManager INSTANCE;

  private CsisResource main = CsisResource.CSIS;

  private CsisConnectionManager() {}

  public static CsisConnectionManager getInstance() {
    if (INSTANCE == null) {
      synchronized (CsisConnectionManager.class) {
        if (INSTANCE == null)
          INSTANCE = new CsisConnectionManager();
      }
    }

    return INSTANCE;
  }

  public SqlSessionFactory getMain() {
    return main.getSessionFactory();
  }

  public void setMain(CsisResource main) {
    this.main = main;
  }

}
