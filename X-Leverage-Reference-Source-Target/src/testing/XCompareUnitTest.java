package testing;

import static org.junit.Assert.assertTrue;

import java.io.File;
import java.net.URI;
import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.desktoptool.xcompare.FileConverter;

public class XCompareUnitTest {

	private List<String> toclean;
	private com.desktoptool.xcompare.Configuration config;
	
	@Before 
	public void setUp() throws Exception {
		System.out.println("setup");
		
		toclean = new ArrayList<String>();
		File configfile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/XCompare_config.ini").getFile()).getPath());
		config = new com.desktoptool.xcompare.Configuration();
		config.loadConfiguration(configfile.getAbsolutePath());
	}
	
	@Test
	public void convert_MSWord_Get_Txml() throws Exception{

		System.out.println("running test: \"" + "convert_MSWord_Get_Txml" + "\"");
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset2/word.docx").getFile()).getPath());
		FileConverter.convertWordToTxml(sourcefile.getAbsolutePath(), config.getConfigurationFile(), "EN");
		String report = sourcefile.getParent() + File.separator + sourcefile.getName() + ".txml";
		toclean.add(report);
		assertTrue(new File(report).exists());
	}
	
	@Test
	public void convert_MSExcel_Get_Txml() throws Exception{

		System.out.println("running test: \"" + "convert_MSExcel_Get_Txml" + "\"");
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset2/excel.xlsx").getFile()).getPath());
		FileConverter.convertExcelToTxml(sourcefile.getAbsolutePath(), config.getConfigurationFile(), "EN");
		String report = sourcefile.getParent() + File.separator + sourcefile.getName() + ".txml";
		String excelConfig = sourcefile.getParent() + File.separator + "excelConfig.tmp";
		toclean.add(report);
		toclean.add(excelConfig);
		assertTrue(new File(report).exists());
	}

	@After
	public void tearDown() throws Exception {
		System.out.println("cleanup");
		
		for(String f : toclean){
			org.apache.commons.io.FileUtils.forceDelete(new File(f));
		}
	}
}
