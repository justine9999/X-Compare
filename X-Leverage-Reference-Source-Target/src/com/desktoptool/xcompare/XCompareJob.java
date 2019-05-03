package com.desktoptool.xcompare;

import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.Callable;


public class XCompareJob implements Callable<String> {

	public int index;
	public String sourcelanguage;
	public String targetlanguage;
	public List<File> sourcelistfile;
	public List<File> sourcereferencelistfile;
	public List<File> targetreferencelistfile;
	private String tempfolder;
	private String configurationfile;
	private List<String> progresses;

	
	public XCompareJob(int index, String sourcelanguage, String targetlanguage, List<File> sourcelistfile,
			List<File> sourcereferencelistfile, List<File> targetreferencelistfile, String tempfolder,
			String configurationfile, List<String> progresses) {
		super();
		this.index = index;
		this.sourcelanguage = sourcelanguage;
		this.targetlanguage = targetlanguage;
		this.sourcelistfile = sourcelistfile;
		this.sourcereferencelistfile = sourcereferencelistfile;
		this.targetreferencelistfile = targetreferencelistfile;
		this.tempfolder = tempfolder;
		this.configurationfile = configurationfile;
		this.progresses = progresses;
	}


	public String call() {
		
		StringBuffer sb = new StringBuffer();
		
		if(sourcelistfile.size() == 0){
			sb.append("[need at least 1 valid source file to run]");
			return sb.toString();
		}
		
		String sourcefileDirectory = sourcelistfile.get(0).getParent();
		
		if(!new File(tempfolder).exists()){
			new File(tempfolder).mkdir();
		}
		try {
			org.apache.commons.io.FileUtils.cleanDirectory(new File(tempfolder));
		} catch (IOException ex) {
			sb.append("cannot clean temporary folder, maybe it's open or being used by other process!\n");	
			sb.append(ex.getStackTrace());
			return sb.toString();
		}
		
		
		try {
			Thread.sleep(3000);
		} catch (InterruptedException ex) {
			progresses.set(this.index, "F");
			StringWriter sw = new StringWriter();
			PrintWriter pw = new PrintWriter(sw);
			ex.printStackTrace(pw);
			sb.append(sw.toString());
		}
		
		progresses.set(this.index, "C");

		try {
			
			sb.append("[converting source reference files]\n");
			String tempsourcereferencefolder = tempfolder + File.separator + "ref";
			new File(tempsourcereferencefolder).mkdir();
			sourcereferencelistfile = FileUtils.convertSourceFilesToTxmls(sourcereferencelistfile, configurationfile, sourcelanguage, tempsourcereferencefolder, sb);
			
			sb.append("[converting target reference files]\n");
			String temptargetreferencefolder = tempfolder + File.separator + "tref";
			new File(temptargetreferencefolder).mkdir();
			targetreferencelistfile = FileUtils.convertSourceFilesToTxmls(targetreferencelistfile, configurationfile, targetlanguage, temptargetreferencefolder, sb);
			
			
			sb.append("[XCompare job stating]\n");
			
			Configuration config = new Configuration();
			if(!configurationfile.isEmpty()) {
				config.loadConfiguration(configurationfile);
			}
			
			Analyzer analyzer = new Analyzer(config, this.index, this.progresses, sb);
			analyzer.setTempFolder(tempfolder);
			
			HashMap<String, List<String>> sourcefilesMap = new HashMap<String, List<String>>();
			HashMap<String, List<String>> sourceReferencefilesMap = new HashMap<String, List<String>>();
			HashMap<String, List<String>> targetReferencefilesMap = new HashMap<String, List<String>>();
			
			sb.append("[collecting file information]\n");
			
			String lang_pair_code = sourcelanguage + "_" + targetlanguage;
			
			for(File file : sourcelistfile) {
				
				String ext = FileUtils.getExtension(file.getName());
				if(ext.toLowerCase().equals("txml")) {
					
					String langCode = FileUtils.getTxmlLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = lang_pair_code;
					if(!sourcefilesMap.containsKey(langCode)){
						sourcefilesMap.put(langCode, new ArrayList<String>());
					}
					sourcefilesMap.get(langCode).add(file.getAbsolutePath());
					
				} else if(ext.toLowerCase().equals("txlf")) {
					
					String langCode = FileUtils.getTxlfLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = lang_pair_code;
					if(!sourcefilesMap.containsKey(langCode)){
						sourcefilesMap.put(langCode, new ArrayList<String>());
					}
					sourcefilesMap.get(langCode).add(file.getAbsolutePath());
				}
			}
			
			sb.append("[source files ready to be run]\n");
			for(String lang : sourcefilesMap.keySet()){
				sb.append(lang + ":" + "\n");
				for(String file : sourcefilesMap.get(lang)){
					sb.append(new File(file).getName() + "\n");
				}
			}
			
			
			String default_src_lang_code = sourcelanguage;
			
			for(File file : sourcereferencelistfile) {
				
				String ext = FileUtils.getExtension(file.getName());
				if(ext.toLowerCase().equals("txml")) {
					
					String langCode = FileUtils.getTxmlLanguageCode(file.getAbsolutePath()).split("_")[0];
					if(!config.isForce_language_match()) langCode = default_src_lang_code;
					if(!sourceReferencefilesMap.containsKey(langCode)){
						sourceReferencefilesMap.put(langCode, new ArrayList<String>());
					}
					sourceReferencefilesMap.get(langCode).add(file.getAbsolutePath());
					
				} else if(ext.toLowerCase().equals("txlf")) {
					
					String langCode = FileUtils.getTxlfLanguageCode(file.getAbsolutePath());
					if(!config.isForce_language_match()) langCode = default_src_lang_code;
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
			
			sb.append("[source reference files ready to be run]\n");
			for(String lang : sourceReferencefilesMap.keySet()){
				sb.append(lang + ":" + "\n");
				for(String file : sourceReferencefilesMap.get(lang)){
					sb.append(new File(file).getName() + "\n");
				}
			}
			
			String default_trg_lang_code = targetlanguage;
			
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
			
			sb.append("[target reference files ready to be run]\n");
			for(String lang : targetReferencefilesMap.keySet()){
				sb.append(lang + ":" + "\n");
				for(String file : targetReferencefilesMap.get(lang)){
					sb.append(new File(file).getName() + "\n");
				}
			}
			
			sb.append("[analyzing files]\n");
			
			for(String langCode : sourcefilesMap.keySet()) {

				String srclangCode = langCode.split("_")[0];
				String trglangCode = langCode.split("_")[1];

				if(sourceReferencefilesMap.containsKey(srclangCode) || sourceReferencefilesMap.containsKey("xx-xx")) {
					
					analyzer.analyze(sourcefilesMap.get(langCode), sourceReferencefilesMap.getOrDefault(srclangCode, sourceReferencefilesMap.get("xx-xx")), targetReferencefilesMap.getOrDefault(trglangCode, new ArrayList<String>()), langCode, sourcefileDirectory);
				}
			}
			
			sb.append("[XCompare job done]");

			
		} catch(Exception ex){
			
			progresses.set(this.index, "F");
			try {
				Thread.sleep(3000);
			} catch (InterruptedException e) {
				StringWriter sw = new StringWriter();
				PrintWriter pw = new PrintWriter(sw);
				ex.printStackTrace(pw);
				sb.append(sw.toString());
			}
			StringWriter sw = new StringWriter();
			PrintWriter pw = new PrintWriter(sw);
			ex.printStackTrace(pw);
			sb.append(sw.toString());
		}

		return sb.toString();
	}
}
