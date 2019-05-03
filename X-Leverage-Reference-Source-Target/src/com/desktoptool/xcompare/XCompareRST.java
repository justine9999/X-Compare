package com.desktoptool.xcompare;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.function.Function;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.gs4tr.foundation.core.commandline.Arguments;


public class XCompareRST {
	
	public static void main(String[] argv)
	{
		System.exit(main2(argv));
	}

	public static int main2(String[] args) {
		
		String sourcefilePath = "";
		String sourceReferencefilePath = "";
		String targetReferencefilePath = "";
		String configurationfile = "";
		
		
		Arguments getopt = new Arguments();
	      
		getopt.setUsage(new String[] { "Usage: java XCompareRST [-s source files Path] [-r source reference files Path] [-t target reference files Path] [-c configuration file]", "", "X-Compare Reference Soure Target" });
		getopt.parseArgumentTokens(args, new char[] { 's' , 'r' , 't'  , 'c'});
		int c;
		while ((c = getopt.getArguments()) != -1) {
			switch (c)
			{
				case 115: 
					sourcefilePath = getopt.getStringParameter();
					break;
				case 114: 
					sourceReferencefilePath = getopt.getStringParameter();
					break;
				case 116: 
					targetReferencefilePath = getopt.getStringParameter();
					break;
				case 99: 
					configurationfile = getopt.getStringParameter();
					if (!new File(configurationfile).exists()) {
						System.err.println("invalid argument for -c");
						System.out.println();
						pressAnyKeyToContinue();
						return 1;
					}
					break;
			}
		}

		Function<String, List<File>> getlistfile = s -> {
				String[] paths = s.split(";");
				List<File> files = new ArrayList<File>();
				for(String path: paths){
					File file = new File(path);
					if(file.exists() && file.isFile()){
						files.add(file);
					}
				}
				return files;
			};
			
		List<File> sourcelistfile = getlistfile.apply(sourcefilePath);
		List<File> sourcereferencelistfile = getlistfile.apply(sourceReferencefilePath);
		List<File> targetreferencelistfile = getlistfile.apply(targetReferencefilePath);
		
		String sourcefileP = "N:\\j\\folders\\jx'\\XCompareRST\\file\\src";
		String sourceReferencefileP = "N:\\j\\folders\\jx'\\XCompareRST\\file\\ref";
		String targetReferencefileP = "N:\\j\\folders\\jx'\\XCompareRST\\file\\tref";
		configurationfile = "N:\\j\\folders\\jx'\\XCompareRST\\XCompare_config.ini";
		Function<String, List<File>> gatherlistfile = s -> {
			List<File> files = new ArrayList<File>();
			for(File file : new File(s).listFiles()){
				if(file.isFile()) files.add(file);
			}
			return files;
		};
		
		sourcelistfile = gatherlistfile.apply(sourcefileP);
		sourcereferencelistfile = gatherlistfile.apply(sourceReferencefileP);
		targetreferencelistfile = gatherlistfile.apply(targetReferencefileP);
		
		if(sourcelistfile.size() == 0){
			System.err.println("[need at least 1 valid source file to run]");
			System.out.println();
			pressAnyKeyToContinue();
			return 1;
		}
		
		String sourcefileDirectory = sourcelistfile.get(0).getParent();
		
		String tempfolder = System.getProperty("java.io.tmpdir") + File.separator + "x-compare";
		if(!new File(tempfolder).exists()){
			new File(tempfolder).mkdir();
		}
		try {
			org.apache.commons.io.FileUtils.cleanDirectory(new File(tempfolder));
		} catch (IOException e) {
			System.err.println("cannot clean temporary folder, maybe it's open or being used by other process!");	
			e.printStackTrace();
			System.out.println();
			pressAnyKeyToContinue();
			return 1;
		}
		
		
		String sourcelanguage = FileUtils.getTxmlLanguageCode(sourcelistfile.get(0).getAbsolutePath()).split("_")[0];
		
		ApplyAsposeLicenseToWord();
		BasicConfigurator.configure();
		Logger.getRootLogger().setLevel(Level.OFF);
		
		try {
			
			System.out.println("[converting source reference files]");
			String tempsourcereferencefolder = tempfolder + File.separator + "ref";
			new File(tempsourcereferencefolder).mkdir();
			sourcereferencelistfile = FileUtils.convertSourceFilesToTxmls(sourcereferencelistfile, configurationfile, sourcelanguage, tempsourcereferencefolder, new StringBuffer());
			
			System.out.println();
			System.out.println("[converting target reference files]");
			String temptargetreferencefolder = tempfolder + File.separator + "tref";
			new File(temptargetreferencefolder).mkdir();
			targetreferencelistfile = FileUtils.convertSourceFilesToTxmls(targetreferencelistfile, configurationfile, sourcelanguage, temptargetreferencefolder, new StringBuffer());
			
			
			System.out.println();
			System.out.println("[XCompare job stating]");
			
			Configuration config = new Configuration();
			if(!configurationfile.isEmpty()) {
				config.loadConfiguration(configurationfile);
			}
			
			Analyzer analyzer = new Analyzer(config, 0, new ArrayList<String>(), new StringBuffer());
			analyzer.setTempFolder(tempfolder);
			
			HashMap<String, List<String>> sourcefilesMap = new HashMap<String, List<String>>();
			HashMap<String, List<String>> sourceReferencefilesMap = new HashMap<String, List<String>>();
			HashMap<String, List<String>> targetReferencefilesMap = new HashMap<String, List<String>>();
			
			System.out.println();
			System.out.println("[collecting file information]");
			
			String default_lang_code = FileUtils.getTxmlLanguageCode(sourcelistfile.get(0).getAbsolutePath());
			
			for(File file : sourcelistfile) {
				
				String ext = FileUtils.getExtension(file.getName());
				if(ext.toLowerCase().equals("txml")) {
					
					String langCode = FileUtils.getTxmlLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = default_lang_code;
					if(!sourcefilesMap.containsKey(langCode)){
						sourcefilesMap.put(langCode, new ArrayList<String>());
					}
					sourcefilesMap.get(langCode).add(file.getAbsolutePath());
					
				} else if(ext.toLowerCase().equals("txlf")) {
					
					String langCode = FileUtils.getTxlfLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = default_lang_code;
					if(!sourcefilesMap.containsKey(langCode)){
						sourcefilesMap.put(langCode, new ArrayList<String>());
					}
					sourcefilesMap.get(langCode).add(file.getAbsolutePath());
				}
			}
			
			System.out.println();
			System.out.println("[source files ready to be run]");
			for(String lang : sourcefilesMap.keySet()){
				System.out.println(lang + ":");
				for(String file : sourcefilesMap.get(lang)){
					System.out.println(new File(file).getName());
				}
			}
			
			
			for(File file : sourcereferencelistfile) {
				
				String ext = FileUtils.getExtension(file.getName());
				if(ext.toLowerCase().equals("txml")) {
					
					String langCode = FileUtils.getTxmlLanguageCode(file.getAbsolutePath()).split("_")[0];
					if(!config.isForce_language_match()) langCode = default_lang_code.split("_")[0];
					if(!sourceReferencefilesMap.containsKey(langCode)){
						sourceReferencefilesMap.put(langCode, new ArrayList<String>());
					}
					sourceReferencefilesMap.get(langCode).add(file.getAbsolutePath());
					
				} else if(ext.toLowerCase().equals("txlf")) {
					
					String langCode = FileUtils.getTxlfLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = default_lang_code.split("_")[0];
					if(!sourceReferencefilesMap.containsKey(langCode)){
						sourceReferencefilesMap.put(langCode, new ArrayList<String>());
					}
					sourceReferencefilesMap.get(langCode).add(file.getAbsolutePath());
				} else if(ext.toLowerCase().equals("xlsx")) {

					sourceReferencefilesMap = new HashMap<String, List<String>>();
					sourceReferencefilesMap.put("xx-xx", new ArrayList<String>());
					sourceReferencefilesMap.get("xx-xx").add(file.getAbsolutePath());
					break;
				}
			}
			
			System.out.println();
			System.out.println("[source reference files ready to be run]");
			for(String lang : sourceReferencefilesMap.keySet()){
				System.out.println(lang + ":");
				for(String file : sourceReferencefilesMap.get(lang)){
					System.out.println(new File(file).getName());
				}
			}
			
			String default_trg_lang_code = default_lang_code.split("_")[1];
			
			for(File file : targetreferencelistfile) {
				
				String ext = FileUtils.getExtension(file.getName());
				if(ext.toLowerCase().equals("txml")) {
					
					String langCode = FileUtils.getTxmlLanguageCode(file.getAbsolutePath()).split("_")[0];
					if(!config.isForce_language_match()) langCode = default_trg_lang_code;
					if(!targetReferencefilesMap.containsKey(langCode)){
						targetReferencefilesMap.put(langCode, new ArrayList<String>());
					}
					targetReferencefilesMap.get(langCode).add(file.getAbsolutePath());
					
				} else if(ext.toLowerCase().equals("txlf")) {
					
					String langCode = FileUtils.getTxlfLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = default_trg_lang_code;
					if(!targetReferencefilesMap.containsKey(langCode)){
						targetReferencefilesMap.put(langCode, new ArrayList<String>());
					}
					targetReferencefilesMap.get(langCode).add(file.getAbsolutePath());
				}
			}
			
			System.out.println();
			System.out.println("[target reference files ready to be run]");
			for(String lang : targetReferencefilesMap.keySet()){
				System.out.println(lang + ":");
				for(String file : targetReferencefilesMap.get(lang)){
					System.out.println(new File(file).getName());
				}
			}
			
			System.out.println();
			System.out.println("[analyzing files]");
			
			for(String langCode : sourcefilesMap.keySet()) {

				String srclangCode = langCode.split("_")[0];
				String trglangCode = langCode.split("_")[1];

				if(sourceReferencefilesMap.containsKey(srclangCode) || sourceReferencefilesMap.containsKey("xx-xx")) {
					
					analyzer.analyze(sourcefilesMap.get(langCode), sourceReferencefilesMap.getOrDefault(srclangCode, sourceReferencefilesMap.get("xx-xx")), targetReferencefilesMap.getOrDefault(trglangCode, new ArrayList<String>()), langCode, sourcefileDirectory);
				}
			}
			
			System.out.println();
			System.out.println("[XCompare job done]");

			
		} catch(Exception ex){
			
			ex.printStackTrace();
		}
		
		System.out.println();
		pressAnyKeyToContinue();
		
		return 0;
	}
	
	public static void ApplyAsposeLicenseToWord()
	{
		com.aspose.cells.License license= new com.aspose.cells.License();
		try {
			System.out.println();
			System.out.println("[applying aspose license]");
			InputStream licence = XCompareRST.class.getClassLoader().getResourceAsStream("properties/aspose.license");
			license.setLicense(licence);
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void pressAnyKeyToContinue()
	{ 
        System.out.println("[Press Enter key to close]");
        try
        {
            System.in.read();
        }  
        catch(Exception e)
        {}  
	}

}
