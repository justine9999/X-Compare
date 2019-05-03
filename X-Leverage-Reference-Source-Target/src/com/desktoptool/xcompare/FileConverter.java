package com.desktoptool.xcompare;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.configuration.Configuration;
import org.apache.commons.configuration.PropertiesConfiguration;
import org.gs4tr.filters.msoffice.ConvertDOC;
import org.gs4tr.filters.msoffice.ConvertXLS;
import org.gs4tr.foundation.locale.Locale;


public class FileConverter {

	public static String convertWordToTxml(String source, String configfiledir, String sourcelanguage) throws Exception {
		
		List<String> sources = new ArrayList<String>();
		sources.add(source);
		Locale locale = Locale.makeLocale(sourcelanguage);
		Configuration config = new PropertiesConfiguration(configfiledir);

		ConvertDOC converter = new ConvertDOC();
		converter.setConfiguration(config);
		converter.setIgnoreSuccessfullConversion(true);
		converter.convertFiles(sources, locale);
		
		String txml = source + ".txml";
		if(!new File(txml).exists()) {
			throw new FileNotFoundException();
		}
		
		return txml;
	}
	
	public static String convertExcelToTxml(String source, String sourcelanguage, String configfiledir) throws Exception {
		
		List<String> sources = new ArrayList<String>();
		sources.add(source);
		Locale locale = Locale.makeLocale(sourcelanguage);
		Configuration config = new PropertiesConfiguration(configfiledir);

		String excelConfigFile = FileUtils.makeExcelConfigFile(new File(source).getParent());
		ConvertXLS converter = new ConvertXLS();
		converter.setConfiguration(config);
		converter.setIgnoreSuccessfullConversion(true);
		converter.convertFiles(sources, locale, excelConfigFile);
		
		String txml = source + ".txml";
		if(!new File(txml).exists()) {
			throw new FileNotFoundException();
		}
		
		return txml;
	}
}
