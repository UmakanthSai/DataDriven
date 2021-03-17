package automation;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.util.SystemOutLogger;

public class SampleTest {

	public static void main(String[] args) throws IOException {
		
		Test m = new Test();
		
		ArrayList d =m.ExcelData("DataDriven", "TestCases", "sample1");
		
		for(int i=0; i<d.size(); i++) {
			System.out.println(d.get(i));
		}
 
	}

}
