package testing;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.desktoptool.xcompare.Analyzer;
import com.desktoptool.xcompare.FileUtils;

public class XCompareTest {

	private com.desktoptool.xcompare.Configuration config;
	private Analyzer analyzer;
	private String tempfolder;
	private String sourcefolder;
	private String sourcereferencefolder;
	private String targetreferencefolder;
	
	@Before 
	public void setUp() throws Exception {
		System.out.println("setup");
		
		File configfile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/XCompare_config.ini").getFile()).getPath());
		config = new com.desktoptool.xcompare.Configuration();
		config.loadConfiguration(configfile.getAbsolutePath());
		List<String> progresses = new ArrayList<String>();
		progresses.add("S");
		analyzer = new Analyzer(config, 0, progresses, new StringBuffer());
	}
	
	@Test
	public void run_Only_Source_And_Source_Reference_Get_Excel_Report() throws Exception{
		
		tempfolder = System.getProperty("java.io.tmpdir") + File.separator + "x-compare";
		if(!new File(tempfolder).exists()){
			new File(tempfolder).mkdir();
		}
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/src/src.txml").getFile()).getPath());
		sourcefolder = sourcefile.getParent();
		List<String> sourcefiles = new ArrayList<String>();
		sourcefiles.add(sourcefile.getAbsolutePath());
		
		File sourcereferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/ref/ref.docx").getFile()).getPath());
		sourcereferencefolder = sourcereferencefile.getParent();
		List<File> sourceReferencefiles = new ArrayList<File>();
		sourceReferencefiles.add(sourcereferencefile);
		
		String tempsourcereferencefolder = tempfolder + File.separator + "ref";
		new File(tempsourcereferencefolder).mkdir();
		List<File> sourcereferencefilesFile = FileUtils.convertSourceFilesToTxmls(sourceReferencefiles, config.getConfigurationFile(), "EN", tempsourcereferencefolder, new StringBuffer());

		
		List<String> targetReferencefiles = new ArrayList<String>();
		
		String languagePairCode = "EN_DA";
		
		String outputDirectory = sourcefile.getParent();

		analyzer.analyze(sourcefiles, sourcereferencefilesFile.stream().map(f -> f.getAbsolutePath()).collect(Collectors.toList()), targetReferencefiles, languagePairCode, outputDirectory);
		
		String report = analyzer.getTargetReport();
		assertTrue(new File(report).exists());
	}
	
	//@Test
	//public void run_Source_And_Source_Reference_And_Target_Reference_With_AutoAligner_Get_Excel_Report() {
		
	//}
	
	//@Test
	//public void run_Source_And_Source_Reference_And_Target_Reference_Without_AutoAligner_Get_Excel_Report() {
		
	//}

	@After
	public void tearDown() throws Exception {
		System.out.println("cleanup");
		org.apache.commons.io.FileUtils.cleanDirectory(new File(tempfolder));
		if(sourcefolder != null){
			for(File f : new File(sourcefolder).listFiles()){
				if(f.getName().startsWith("X-Compare_")){
					org.apache.commons.io.FileUtils.forceDelete(f);
				}
			}
		}
	}
}
