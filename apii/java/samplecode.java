import java.net.*;
import java.io.*;

public class samplecode 
{
  public samplecode()
  {
  }

  public static void main(String[] args) {
      String sMsg = "test sms from api";
      String sURL = "http://sample.onewaysms.com.au:xxxx/api.aspx?apiusername=xyz&apipassword=xyz&mobileno=6141234567&senderid=onewaysms&languagetype=1&message=" + URLEncoder.encode(sMsg);
      String result = "";  
      HttpURLConnection conn = null;
      try  {
          URL url = new URL(sURL);
          conn = (HttpURLConnection)url.openConnection();          
          conn.setDoOutput(false);                  
          conn.setRequestMethod("GET");          
          conn.connect();
          int iResponseCode = conn.getResponseCode();
          if ( iResponseCode == 200 ) {
            BufferedReader oIn = new BufferedReader(new InputStreamReader(conn.getInputStream())); 
            String sInputLine = "";
            String sResult = "";
            while ((sInputLine = oIn.readLine()) != null) {
              sResult = sResult + sInputLine;
            }
            if (Long.parseLong(sResult) > 0) 
            {
              System.out.println("success - MT ID : " + sResult);       
            }
            else 
            {
              System.out.println("fail - Error code : " + sResult);       
            }
          }
          else {
            System.out.println("fail ");        
          }
      }
      catch (Exception e){ 
        e.printStackTrace();
      }
      finally {
        if (conn != null) {
          conn.disconnect();
        }
      }  
    }
}