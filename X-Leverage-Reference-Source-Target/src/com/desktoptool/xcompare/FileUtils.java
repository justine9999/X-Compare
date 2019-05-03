package com.desktoptool.xcompare;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.dom4j.Attribute;
import org.dom4j.Element;
import org.gs4tr.foundation.xml.XmlParser;
import org.w3c.dom.Document;

public class FileUtils {
	
	public static String getExtension(String filepath) {
		return filepath.substring(filepath.lastIndexOf('.')+1, filepath.length());
	}
	
	public static String getNameWithoutExtension(String filepath) {
		String filename = new File(filepath).getName();
		return filename.substring(0, filename.lastIndexOf('.'));
	}
	
	public static String getTxmlLanguageCode(String filepath) {
		
		org.dom4j.Document document = XmlParser.parseXmlFile(filepath);
		Element root = document.getRootElement();
		String sourceCode = root.attributeValue("locale");
		String targetCode = root.attributeValue("targetlocale");
		return sourceCode.toUpperCase() + "_" + (targetCode == null?"LL-CC":targetCode.toUpperCase());
	}
	
	public static String getTxlfLanguageCode(String filepath) {
		
		org.dom4j.Document document = XmlParser.parseXmlFile(filepath);
		Element root = document.getRootElement();
		Element file = root.element("file");
		String sourceCode = file.attributeValue("source-language");
		String targetCode = file.attributeValue("target-language");
		return sourceCode.toUpperCase() + "_" + targetCode.toUpperCase();
	}
	
	public static void readSegments(String file, List<XSegment> segments, boolean read_source, Configuration config) {
		
		String ext = FileUtils.getExtension(file);
		if(ext.toLowerCase().equals("txml")) {
			readSegmentsTxml(file, segments, read_source, config);
		} else if(ext.toLowerCase().equals("txlf")) {
			readSegmentsTxlf(file, segments, read_source, config);
		}
	}
	
	@SuppressWarnings("rawtypes")
	private static void readSegmentsTxml(String file, List<XSegment> listsegments, boolean read_source, Configuration config) {
		
		org.dom4j.Document document = XmlParser.parseXmlFile(file);
		Element root = document.getRootElement();
		String sourceLanguageCode = root.attributeValue("locale").substring(0, 2).toLowerCase();
		String targetLanguageCode = root.attributeValue("targetlocale") != null?root.attributeValue("targetlocale").substring(0, 2).toLowerCase():"EN";
		List segments = root.selectNodes("//*[name() = 'segment']");
		
		String filename = new File(file).getName();
		for(int i = 0; i < segments.size(); i++) {
			
			Element segment = (Element)segments.get(i);
			if(read_source){
				Element source = segment.element("source");
				if(source != null) {
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), source.getText(), NormalizationUtils.normalize(source.getText(), config, sourceLanguageCode));
					listsegments.add(xsegment);
				}else{
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), "", "");
					listsegments.add(xsegment);
				}
			}else{
				Element target = segment.element("target");
				if(target != null) {
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), target.getText(), NormalizationUtils.normalize(target.getText(), config, targetLanguageCode));
					listsegments.add(xsegment);
				}else{
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), "", "");
					listsegments.add(xsegment);
				}
			}
		}
	}
	
	@SuppressWarnings("rawtypes")
	private static void readSegmentsTxlf(String file, List<XSegment> listsegments, boolean read_source, Configuration config) {
		
		org.dom4j.Document document = XmlParser.parseXmlFile(file);
		Element root = document.getRootElement();
		String sourceLanguageCode = root.attributeValue("source-language").substring(0, 2);
		String targetLanguageCode = root.attributeValue("target-language").substring(0, 2);
		List transunits = root.selectNodes("//*[name() = 'trans-unit']");
		
		String filename = new File(file).getName();
		for(int i = 0; i < transunits.size(); i++) {
			
			Element transunit = (Element)transunits.get(i);
			if(read_source){
				Element source = transunit.element("source");
				if(source != null) {
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), source.getText(), NormalizationUtils.normalize(source.getText(), config, sourceLanguageCode));
					listsegments.add(xsegment);
				}else{
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), "", "");
					listsegments.add(xsegment);
				}
			}else{
				Element target = transunit.element("target");
				if(target != null) {
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), target.getText(), NormalizationUtils.normalize(target.getText(), config, targetLanguageCode));
					listsegments.add(xsegment);
				}else{
					XSegment xsegment = new XSegment(filename, Integer.toString(i+1), "", "");
					listsegments.add(xsegment);
				}
			}
		}
	}
	
	@SuppressWarnings("rawtypes")
	public static void populateTargetsByNotes(String file) throws IOException {
		
		org.dom4j.Document document = XmlParser.parseXmlFile(file);
		Element root = document.getRootElement();
		List segments = root.selectNodes("//*[name() = 'segment']");
		
		for(int i = 0; i < segments.size(); i++) {
			
			Element segment = (Element)segments.get(i);
			Element target = segment.element("target");
			if(target == null){
				target = segment.addElement("target");
			}
			
			Element comments = segment.element("comments");
			if(comments == null) continue;
			Element comment = comments.element("comment");
			if(comment == null) continue;
			
			target.setText(comment.getText());
		}

		OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(file)), "UTF8");
	    document.write(writer);
	    writer.close();
	}
	
	@SuppressWarnings("rawtypes")
	public static void populateReportNotesToNotes(String file, HashMap<String, String> notes, HashMap<String, String> conclusions) throws IOException {
		
		String ext = getExtension(new File(file).getName());
		int idx = 0;
		if(ext.toLowerCase().equals("txml")){
			
			org.dom4j.Document document = XmlParser.parseXmlFile(file);
			Element root = document.getRootElement();
			List segments = root.selectNodes("//*[name() = 'segment']");
			
			for(int i = 0; i < segments.size(); i++) {
				idx++;
				Element segment = (Element)segments.get(i);
				Element comments = segment.element("comments");
				if(comments == null){
					comments = segment.addElement("comments");
				}
				Element comment = comments.element("comment");
				if(comment == null){
					comment = comments.addElement("comment");
					comment.addAttribute("creationid", "XCompare");
					comment.addAttribute("creationdate", new SimpleDateFormat("yyyyMMddTHHmmssZ").format(new Date()));
					comment.addAttribute("type", "text");
				}
				String idstr = Integer.toString(idx);
				if(notes.containsKey(idstr)){
					comment.setText(notes.get(idstr));
				}else{
					switch(conclusions.get(idstr)){
						case "0":
							comment.setText("New");
							break;
						case "100-":
							comment.setText("Source revision");
							break;
						case "100":
							comment.setText("Exact");
							break;
						default:
							comment.setText("Source revision");
							break;
					}
				}
			}

			OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(file)), "UTF8");
		    document.write(writer);
		    writer.close();
		}
	}
	
	@SuppressWarnings({ "rawtypes" })
	public static void populateMatchConclusionToNotes(String file, int[] conclusion) throws IOException{
		
		String ext = getExtension(new File(file).getName());
		if(ext.toLowerCase().equals("txml")){
			
			org.dom4j.Document document = XmlParser.parseXmlFile(file);
			Element root = document.getRootElement();
			List segments = root.selectNodes("//*[name() = 'segment']");
			
			for(int i = 0; i < segments.size(); i++) {
				
				Element segment = (Element)segments.get(i);
				Element comments = segment.element("comments");
				if(comments != null) segment.remove(comments);
				comments = segment.addElement("comments");
				Element comment = comments.addElement("comment");
				comment.addAttribute("creationid", "XCompare");
				comment.addAttribute("creationdate", new SimpleDateFormat("yyyyMMdd'T'HHmmss'Z'").format(new Date()));
				comment.addAttribute("type", "text");
				
				switch(conclusion[i]){
					case 0:
						comment.setText("New");
						break;
					case 1:
						comment.setText("Source revision");
						break;
					case 2:
						comment.setText("Source revision");
						break;
					case 3:
						comment.setText("Exact");
						break;
					default:
				}
			}

			OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(file)), "UTF8");
		    document.write(writer);
		    writer.close();
		}
	}
	
	@SuppressWarnings("rawtypes")
	public static void populateAlignedOldTargets(String file, String[] targets, int[] scores, Configuration config) throws IOException{
		
		String ext = getExtension(new File(file).getName());
		if(ext.toLowerCase().equals("txml")){
			
			org.dom4j.Document document = XmlParser.parseXmlFile(file);
			Element root = document.getRootElement();
			List segments = root.selectNodes("//*[name() = 'segment']");
			
			for(int i = 0; i < segments.size(); i++) {
				
				Element segment = (Element)segments.get(i);
				if(segment.attribute("modified") != null){
					segment.remove(segment.attribute("modified"));
				}
				Element target = (Element)segment.element("target");
				if(target != null){
					segment.remove(target);
				}
				target = segment.addElement("target");
				target.addAttribute("score", Integer.toString(scores[i]));
				if(scores[i] >= config.getThreshold_noraml()){
					target.addAttribute("secondary", "true");
				}
				
				target.setText(targets[i]);
			}

			OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(file)), "UTF8");
		    document.write(writer);
		    writer.close();
		}
	}
	
	public static List<File> convertSourceFilesToTxmls(List<File> sources, String configfiledir, String sourcelanguage, String tempfolder, StringBuffer sb) throws Exception{
		
		List<File> txmls = new ArrayList<File>();
		for(File file : sources){
			String ext = FileUtils.getExtension(file.getName());
			if(ext.toLowerCase().equals("txml")){
				txmls.add(file);
			}else if(ext.toLowerCase().equals("doc") || ext.toLowerCase().equals("docx")){
				
				sb.append("=> " + file.getName()+"\n");
				File filecopy = new File(tempfolder + File.separator + file.getName());
				org.apache.commons.io.FileUtils.copyFile(file, filecopy);
				String txml = FileConverter.convertWordToTxml(filecopy.getAbsolutePath(), configfiledir, sourcelanguage);
				txmls.add(new File(txml));
				
			}else if(ext.toLowerCase().equals("xls") || ext.toLowerCase().equals("xlsx")){
				
				if(file.getName().toLowerCase().startsWith("x-compare_")){
					txmls.add(file);
					continue;
				}
				
				sb.append("=> " + file.getName()+"\n");
				File filecopy = new File(tempfolder + File.separator + file.getName());
				org.apache.commons.io.FileUtils.copyFile(file, filecopy);
				String txml = FileConverter.convertExcelToTxml(filecopy.getAbsolutePath(), sourcelanguage, configfiledir);
				txmls.add(new File(txml));
				
			}else{
				continue;
			}	
		}
		
		return txmls;
	}

	public static String makeExcelConfigFile(String basefolder)
	{
		String configFileFullPath = basefolder + File.separator + "excelConfig.tmp";
		try
		{
			DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
			DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
      
			Document doc = docBuilder.newDocument();
			org.w3c.dom.Element rootElement = doc.createElement("configuration");
			doc.appendChild(rootElement);
      
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			DOMSource source = new DOMSource(doc);
			StreamResult result = new StreamResult(new File(configFileFullPath));
      
			transformer.transform(source, result);
		}
		catch (ParserConfigurationException pce)
		{
			pce.printStackTrace();
		}
		catch (TransformerException tfe)
		{
			tfe.printStackTrace();
		}
		return configFileFullPath;
	}
	
	public static void fixTxmlTxlfLanguageCodes(String file, String sourcelanguage, String targetlanguage) throws Exception{
		
		org.dom4j.Document document = XmlParser.parseXmlFile(file);
		Element root = document.getRootElement();
		String att_name_src = null;
		String att_name_trg = null;
		String ext = getExtension(file).toLowerCase();
		if(ext.equals("txml")){

			att_name_src = "locale";
			att_name_trg = "targetlocale";
			
		}else if(ext.equals("txlf")){
			
			att_name_src = "source-language";
			att_name_trg = "target-language";
		}
		
		Attribute sourcelocale = null;
		Attribute targetlocale = null;
		sourcelocale = root.attribute(att_name_src);
		targetlocale = root.attribute(att_name_trg);
		if(sourcelocale != null) root.remove(sourcelocale);
		if(targetlocale != null) root.remove(targetlocale);
		root.addAttribute(att_name_src, sourcelanguage);
		root.addAttribute(att_name_trg, targetlanguage);
		
		
		OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(file)), "UTF8");
	    document.write(writer);
	    writer.close();
	}
}
