package testing;

import static org.junit.Assert.*;

import java.io.File;
import java.net.URI;
import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import com.desktoptool.xcompare.Analyzer;

public class XCompareIntegrationTest {

	private com.desktoptool.xcompare.Configuration config;
	private Analyzer analyzer;
	private String tempfolder;
	private String sourcefolder;
	//private String sourcereferencefolder;
	//private String targetreferencefolder;
	//private String targetferencefolder;
	
	@Before 
	public void setUp() throws Exception {
		System.out.println("setup");
		
		File configfile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/XCompare_config.ini").getFile()).getPath());
		config = new com.desktoptool.xcompare.Configuration();
		config.loadConfiguration(configfile.getAbsolutePath());
		List<String> progresses = new ArrayList<String>();
		progresses.add("S");
		analyzer = new Analyzer(config, 0, progresses, new StringBuffer());
		
		tempfolder = System.getProperty("java.io.tmpdir") + File.separator + "x-compare";
		if(!new File(tempfolder).exists()){
			new File(tempfolder).mkdir();
		}
		
		analyzer.setTempFolder(tempfolder);
	}
	
	@Test
	public void run_Only_Source_And_Source_Reference_Get_Excel_Report() throws Exception{

		System.out.println("running test: \"" + "run_Only_Source_And_Source_Reference_Get_Excel_Report" + "\"");
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/src/src.txml").getFile()).getPath());
		sourcefolder = sourcefile.getParent();
		List<String> sourcefiles = new ArrayList<String>();
		sourcefiles.add(sourcefile.getAbsolutePath());

		File sourcereferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/ref/ref.txml").getFile()).getPath());
		List<String> sourceReferencefiles = new ArrayList<String>();
		sourceReferencefiles.add(sourcereferencefile.getAbsolutePath());
		
		List<String> targetReferencefiles = new ArrayList<String>();
		
		String languagePairCode = "EN_DA";
		
		String outputDirectory = sourcefile.getParent();

		analyzer.analyze(sourcefiles, sourceReferencefiles, targetReferencefiles, languagePairCode, outputDirectory);
		
		String report = analyzer.getTargetReport();
		assertTrue(new File(report).exists());
	}
	
	@Test
	public void run_Source_And_Source_Reference_And_Target_Reference_Without_AutoAligner_Get_Excel_Report() throws Exception{

		System.out.println("running test: \"" + "run_Source_And_Source_Reference_And_Target_Reference_Without_AutoAligner_Get_Excel_Report" + "\"");
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/src/src.txml").getFile()).getPath());
		sourcefolder = sourcefile.getParent();
		List<String> sourcefiles = new ArrayList<String>();
		sourcefiles.add(sourcefile.getAbsolutePath());
		
		/*File sourcereferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/ref/ref.docx").getFile()).getPath());
		sourcereferencefolder = sourcereferencefile.getParent();
		List<File> sourceReferencefiles = new ArrayList<File>();
		sourceReferencefiles.add(sourcereferencefile);
		
		String tempsourcereferencefolder = tempfolder + File.separator + "ref";
		new File(tempsourcereferencefolder).mkdir();
		List<File> sourcereferencefilesFile = FileUtils.convertSourceFilesToTxmls(sourceReferencefiles, config.getConfigurationFile(), "EN", tempsourcereferencefolder, new StringBuffer());*/
		
		File sourcereferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/ref/ref.txml").getFile()).getPath());
		List<String> sourceReferencefiles = new ArrayList<String>();
		sourceReferencefiles.add(sourcereferencefile.getAbsolutePath());

		/*File targetreferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/tref/tref.txml").getFile()).getPath());
		targetferencefolder = targetreferencefile.getParent();
		List<File> targetReferencefiles = new ArrayList<File>();
		targetReferencefiles.add(targetreferencefile);
		
		String temptargetferencefolder = tempfolder + File.separator + "tref";
		new File(temptargetferencefolder).mkdir();
		List<File> targetReferencefilesFile = FileUtils.convertSourceFilesToTxmls(targetReferencefiles, config.getConfigurationFile(), "DA", tempsourcereferencefolder, new StringBuffer());*/
		
		File targetreferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/tref/tref.txml").getFile()).getPath());
		List<String> targetReferencefiles = new ArrayList<String>();
		targetReferencefiles.add(targetreferencefile.getAbsolutePath());
				
		String languagePairCode = "EN_DA";
		
		String outputDirectory = sourcefile.getParent();

		analyzer.analyze(sourcefiles, sourceReferencefiles, targetReferencefiles, languagePairCode, outputDirectory);
		
		String report = analyzer.getTargetReport();
		assertTrue(new File(report).exists());
	}
	
	@Test
	public void run_Source_And_Source_Reference_And_Target_Reference_With_AutoAligner_Get_Excel_Report() throws Exception{
		
		System.out.println("running test: \"" + "run_Source_And_Source_Reference_And_Target_Reference_With_AutoAligner_Get_Excel_Report" + "\"");
		
		config.setAutoalign_previous_translation(true);
		
		File sourcefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/src/src.txml").getFile()).getPath());
		sourcefolder = sourcefile.getParent();
		List<String> sourcefiles = new ArrayList<String>();
		sourcefiles.add(sourcefile.getAbsolutePath());
		
		File sourcereferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/ref/ref.txml").getFile()).getPath());
		List<String> sourceReferencefiles = new ArrayList<String>();
		sourceReferencefiles.add(sourcereferencefile.getAbsolutePath());
		
		File targetreferencefile = new File(new URI(getClass().getClassLoader().getResource("resource/testdata/fileset1/tref/tref.txml").getFile()).getPath());
		List<String> targetReferencefiles = new ArrayList<String>();
		targetReferencefiles.add(targetreferencefile.getAbsolutePath());
				
		String languagePairCode = "EN_DA";
		
		String outputDirectory = sourcefile.getParent();

		analyzer.analyze(sourcefiles, sourceReferencefiles, targetReferencefiles, languagePairCode, outputDirectory);
		
		String report = analyzer.getTargetReport();
		assertTrue(new File(report).exists());
	}

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
