package readoperations;

import java.net.MalformedURLException;
import java.text.ParseException;

public class App {
    public static void main(String[] args) throws InterruptedException, MalformedURLException, ParseException {
        System.out.println("From App.java");

        ReadOperations readOperations = new ReadOperations();
        readOperations.run();        
    } 
}