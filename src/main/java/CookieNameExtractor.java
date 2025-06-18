import java.util.ArrayList;
import java.util.List;

public class CookieNameExtractor {
  public static List<String> getCookieNames(String cookieString) {
    List<String> cookieNames = new ArrayList<>();
    if (cookieString == null || cookieString.trim().isEmpty()) {
      return cookieNames;
    }

    // Split on semicolon to separate cookies
    String[] cookies = cookieString.split(";");
    for (String cookie : cookies) {
      // Trim to remove leading/trailing whitespace
      cookie = cookie.trim();
      if (cookie.isEmpty()) {
        continue;
      }

      // Find the first '=' to extract the name
      int equalsIndex = cookie.indexOf('=');
      if (equalsIndex != -1) {
        String name = cookie.substring(0, equalsIndex).trim();
        if (!name.isEmpty()) {
          cookieNames.add(name);
        }
      }
    }
    return cookieNames;
  }

  public static void main(String[] args) {
    String cookie =
        "trace_id=6bbc1760e862312b89add29c7e88fc25;"
            + " 57b4dd154cc37901f5ac4735e1128a66=0cb383135cc8c1ddd438ce12754ab60d;"
            + " JSESSIONID=DAC522174310A2E4643E0833B03A2354;"
            + " IDENT-TOKEN=$2a$10$eitFDRsUjCdOIyfOfjXty.4u48UkOZdjYmMdRGJdsecH4FD5ahfPa;"
            + " CSRF-TOKEN=OLVX5_bWd6V6EAQFa8Cyx9JmIsUYCxCAD5Q6cyhrUvo;"
            + " ubiid_access_token_fingerprint=ab64e1d3fdd4be20eb97;"
            + " ubiid_token_access_window=1750326651|oWay5UP9akMvyMrZSpwFRok+WaXfvMq1Eo75Qz8S50c";
    List<String> cookieName = getCookieNames(cookie);
    System.out.println("Cookie Name: " + cookieName); // Output: Cookie Name: oauth2_auth_request
  }
}
