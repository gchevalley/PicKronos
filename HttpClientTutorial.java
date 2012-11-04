import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.URL;
import java.security.KeyManagementException;
import java.security.KeyStore;
import java.security.KeyStoreException;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.security.UnrecoverableKeyException;
import java.security.cert.CertificateException;

import javax.net.ssl.HttpsURLConnection;
import javax.net.ssl.KeyManagerFactory;
import javax.net.ssl.SSLContext;
import javax.net.ssl.SSLPeerUnverifiedException;
import javax.net.ssl.SSLSocketFactory;

import java.security.cert.Certificate;






public class HttpClientTutorial {

    @SuppressWarnings("unused")
    private static javax.net.ssl.SSLSocketFactory getFactory( File pKeyFile, String pKeyPassword ) throws NoSuchAlgorithmException, KeyStoreException, CertificateException, IOException, UnrecoverableKeyException, KeyManagementException  
    {
          KeyManagerFactory keyManagerFactory = KeyManagerFactory.getInstance("SunX509");
          KeyStore keyStore = KeyStore.getInstance("PKCS12");

          InputStream keyInput = new FileInputStream(pKeyFile);
          keyStore.load(keyInput, pKeyPassword.toCharArray());
          keyInput.close();

          keyManagerFactory.init(keyStore, pKeyPassword.toCharArray());

          SSLContext context = SSLContext.getInstance("TLS");
          context.init(keyManagerFactory.getKeyManagers(), null, new SecureRandom());

          return context.getSocketFactory();
    }

       private static void print_https_cert(HttpsURLConnection con){

            if(con!=null){

              try {

            System.out.println("Response Code : " + con.getResponseCode());
            System.out.println("Cipher Suite : " + con.getCipherSuite());
            System.out.println("\n");

            Certificate[] certs = con.getServerCertificates();
            for(Certificate cert : certs){
               System.out.println("Cert Type : " + cert.getType());
               System.out.println("Cert Hash Code : " + cert.hashCode());
               System.out.println("Cert Public Key Algorithm : " + cert.getPublicKey().getAlgorithm());
               System.out.println("Cert Public Key Format : " + cert.getPublicKey().getFormat());
               System.out.println("\n");
            }

            } catch (SSLPeerUnverifiedException e) {
                e.printStackTrace();
            } catch (IOException e){
                e.printStackTrace();
            }

             }

           }

           private static void print_content(HttpsURLConnection con){
            if(con!=null){

            try {

               System.out.println("****** Content of the URL ********");            
               BufferedReader br = 
                new BufferedReader(
                    new InputStreamReader(con.getInputStream()));

               String input;

               while ((input = br.readLine()) != null){
                  System.out.println(input);
               }
               br.close();

            } catch (IOException e) {
               e.printStackTrace();
            }

               }

           }


    public static void main(String[] args) throws IOException, UnrecoverableKeyException, KeyManagementException, NoSuchAlgorithmException, KeyStoreException, CertificateException {
        URL url = new URL("https://bws.bloomberg.com/TmsgServiceSOAP");
        HttpsURLConnection con = (HttpsURLConnection) url.openConnection();
        
        con.setSSLSocketFactory(getFactory(new File("D:\\blp\\data\\tmsg-webservice-csharp\\piccbws.p12"), "ct-84ESh"));

          //dumpl all cert info
         //print_https_cert(con);

         //dump all the content
         //print_content(con);
        
        
        
        String reqXML = new String();
        
        reqXML= "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" xmlns:tmsg=\"http://www.bloomberg.com/services/tmsg\">";
        reqXML = reqXML + "<soapenv:Header/>";
        reqXML = reqXML + "<soapenv:Body>";
        reqXML = reqXML + "<tmsg:TradeIdeaRead>";
        reqXML = reqXML + "<tmsg:Senders>";
        reqXML = reqXML + "<tmsg:Sender>";
        reqXML = reqXML + "<tmsg:Login>jstouff</tmsg:Login>";
        reqXML = reqXML + "</tmsg:Sender>";
        reqXML = reqXML + "</tmsg:Senders>";
        reqXML = reqXML + "</tmsg:TradeIdeaRead>";
        reqXML = reqXML + "</soapenv:Body>";
        reqXML = reqXML + "</soapenv:Envelope>";
     
     
     con.setRequestMethod("POST");
		
		con.setRequestProperty("Content-type", "text/xml; charset=utf-8");
		con.setRequestProperty("SOAPAction", "\"http://www.bloomberg.com/services/tmsg/TradeIdeaRead\"");
		con.setDoOutput(true);
		
		OutputStream reqStream = con.getOutputStream();
		reqStream.write(reqXML.getBytes());
		reqStream.flush();
		
		BufferedReader rd = new BufferedReader(new InputStreamReader(con.getInputStream()));
		String line;
		
		while ((line = rd.readLine()) != null) { System.out.println(line); /*jEdit: print(line); */ }
     
        


    }



}