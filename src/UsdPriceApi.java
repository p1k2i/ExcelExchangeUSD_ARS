import org.json.JSONArray;
import org.json.JSONObject;

import java.io.IOException;
import java.net.URL;
import java.util.Scanner;

public final class UsdPriceApi {
    public static double getPrice(final ConvertType convertType) throws IOException {
        if (convertType.equals(ConvertType.NONE)) return 1d;

        //Instantiating the URL class
        URL url = new URL("https://www.dolarsi.com/api/api.php?type=valoresprincipales");
        //Retrieving the contents of the specified page
        Scanner sc = new Scanner(url.openStream());
        //Instantiating the StringBuffer class to hold the result
        StringBuffer sb = new StringBuffer();
        while(sc.hasNext()) {
            sb.append(sc.next());
            //System.out.println(sc.next());
        }
        //Retrieving the String from the String Buffer object
        String responseResult = sb.toString();
        JSONArray jsonArray = new JSONArray(responseResult);

        for (int i=0; i < jsonArray.length(); i++){
            JSONObject jo = jsonArray.getJSONObject(i);
            JSONObject joCasa = jo.getJSONObject("casa");

            String nombre = joCasa.getString("nombre");

            switch (convertType) {
                case USD_TO_ARS_SELL:
                    if (nombre.replace(" ","").equals("DolarOficial")){
                        return Double.parseDouble(
                                joCasa.getString("venta").replace(',','.')
                        );
                    }
                    break;
                case USD_TO_ARS_BUY:
                    if (nombre.replace(" ","").equals("DolarOficial")){
                        return Double.parseDouble(
                                joCasa.getString("compra").replace(',','.')
                        );
                    }
                    break;
                case ARS_TO_USD_SELL:
                    if (nombre.replace(" ","").equals("DolarOficial")){
                        return 1d/Double.parseDouble(
                                joCasa.getString("venta").replace(',','.')
                        );
                    }
                    break;
                case ARS_TO_USD_BUY:
                    if (nombre.replace(" ","").equals("DolarOficial")){
                        return 1d/Double.parseDouble(
                                joCasa.getString("compra").replace(',','.')
                        );
                    }
                    break;
                case USD_TO_ARS_SELL_BLUE:
                    if (nombre.replace(" ","").equals("DolarBlue")){
                        return Double.parseDouble(
                                joCasa.getString("venta").replace(',','.')
                        );
                    }
                    break;
                case USD_TO_ARS_BUY_BLUE:
                    if (nombre.replace(" ","").equals("DolarBlue")){
                        return Double.parseDouble(
                                joCasa.getString("compra").replace(',','.')
                        );
                    }
                    break;
                case ARS_TO_USD_SELL_BLUE:
                    if (nombre.replace(" ","").equals("DolarBlue")){
                        return 1d/Double.parseDouble(
                                joCasa.getString("venta").replace(',','.')
                        );
                    }
                    break;
                case ARS_TO_USD_BUY_BLUE:
                    if (nombre.replace(" ","").equals("DolarBlue")){
                        return 1d/Double.parseDouble(
                                joCasa.getString("compra").replace(',','.')
                        );
                    }
                    break;
            }
        }

        return 0;
    }
}
