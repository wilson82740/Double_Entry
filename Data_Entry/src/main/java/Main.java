import java.io.IOException;

import dataEntry.DataEntry;

public class Main {
  public static void main(String[] args) {
    DataEntry dataEntry = new DataEntry();
    try {
      dataEntry.runEntey(Long.parseLong(args[0]), Long.parseLong(args[1]));
//      dataEntry.runEntey((long)1007, (long)1007);
    } catch (IOException e) {
      // TODO Auto-generated catch block
      e.printStackTrace();
    }
  }
}
