package testing;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.desktoptool.xcompare.Analyzer;

public class XCompareTest {

	private com.desktoptool.xcompare.Configuration config;
	private Analyzer analyzer;
	
	@Before 
	public void setUp() throws IOException {
		File configfile = new File(getClass().getClassLoader().getResource("resource/testdata/XCompare_config.ini").getFile());
		config = new com.desktoptool.xcompare.Configuration();
		config.loadConfiguration(configfile.getAbsolutePath());
		List<String> progresses = new ArrayList<String>();
		progresses.add("S");
		analyzer = new Analyzer(config, 0, progresses, new StringBuffer());
	}
	
	@Test
	public void run_Only_Source_And_Source_Reference_Get_Excel_Report() {
		
	}
	
	@Test
	public void run_Source_And_Source_Reference_And_Target_Reference_With_AutoAligner_Get_Excel_Report() {
		
	}
	
	@Test
	public void run_Source_And_Source_Reference_And_Target_Reference_Without_AutoAligner_Get_Excel_Report() {
		
	}

	@After
	public void tearDown() {
		System.out.println("after test");
	}
}
