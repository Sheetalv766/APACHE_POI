package excelOperations;
import java.io.IOException;
import java.net.MalformedURLException;
import java.text.ParseException;

public class App {
    public static void main(String[] args) throws InterruptedException, ParseException, IOException {
        System.out.println("From App.java");

        // Uncomment for Milestone 6 activity
        // ReadOperations readOperations = new ReadOperations();
        // readOperations.run();
        
        // Uncomment for Milestone 7 activity
        // WriteOperations writeOperations = new WriteOperations();
        // writeOperations.run();
        
        // Uncomment for Milestone 8 activity
        FormattingOperations formattingOperations = new FormattingOperations();
        formattingOperations.run();        
    } 
}
