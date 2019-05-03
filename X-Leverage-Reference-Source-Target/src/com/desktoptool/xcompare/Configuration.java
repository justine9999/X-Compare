package com.desktoptool.xcompare;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Properties;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class Configuration {
	
	private boolean normalize_whitespace;
	private boolean normalize_case;
	private boolean normalize_englishtonumber;
	private boolean normalize_punctuation;
	private boolean ignore_non_leveragable;
	private boolean ignore_number;
	private boolean show_exact_match;
	private boolean show_partial_exact_match;
	private boolean show_fuzzy_match;
	private boolean show_no_match;
	private boolean show_extra_fuzzy_match_when_exact_match_found;
	private boolean match_merged_reference_segments;
	private boolean force_language_match;
	private boolean autoalign_previous_translation;
	private int threshold_noraml;
	private int threshold_partialmatch;
	private String configurationfile;
	
	
	public Configuration() {
		
		this.normalize_whitespace = true;
		this.normalize_case = false;
		this.normalize_englishtonumber = false;
		this.normalize_punctuation = false;
		this.ignore_non_leveragable = true;
		this.ignore_number = true;
		this.show_exact_match = true;
		this.show_partial_exact_match = true;
		this.show_fuzzy_match = true;
		this.show_no_match = true;
		this.show_extra_fuzzy_match_when_exact_match_found = false;
		this.match_merged_reference_segments = false;
		this.force_language_match = false;
		this.autoalign_previous_translation = false;
		this.threshold_noraml = 75;
		this.threshold_partialmatch = 50;
	}
	
	public void loadConfiguration(String configurationFile) throws IOException {
		
		this.configurationfile = configurationFile;
		FileInputStream in = new FileInputStream(configurationFile);
        Properties props = new Properties();
        props.load(in);
        
        String normalize_whitespace_default = props.getProperty("xcompare.normalization.whitespace");
        if(normalize_whitespace_default != null && !normalize_whitespace_default.trim().isEmpty()) {
        	this.normalize_whitespace = Boolean.parseBoolean(normalize_whitespace_default);
        }
        
        String normalize_case_default = props.getProperty("xcompare.normalization.case");
        if(normalize_case_default != null && !normalize_case_default.trim().isEmpty()) {
        	this.normalize_case = Boolean.parseBoolean(normalize_case_default);
        }
        
        String normalize_englishtonumber_default = props.getProperty("xcompare.normalization.englishtonumber");
        if(normalize_englishtonumber_default != null && !normalize_englishtonumber_default.trim().isEmpty()) {
        	this.normalize_englishtonumber = Boolean.parseBoolean(normalize_englishtonumber_default);
        }
        
        String normalize_punctuation_default = props.getProperty("xcompare.normalization.punctuation");
        if(normalize_punctuation_default != null && !normalize_punctuation_default.trim().isEmpty()) {
        	this.normalize_punctuation = Boolean.parseBoolean(normalize_punctuation_default);
        }
        
        String ignore_non_leveragable_default = props.getProperty("xcompare.ignorenonleveragable");
        if(ignore_non_leveragable_default != null && !ignore_non_leveragable_default.trim().isEmpty()) {
        	this.ignore_non_leveragable = Boolean.parseBoolean(ignore_non_leveragable_default);
        }
        
        String ignore_number_default = props.getProperty("xcompare.ignorenumber");
        if(ignore_number_default != null && !ignore_number_default.trim().isEmpty()) {
        	this.ignore_number = Boolean.parseBoolean(ignore_number_default);
        }
        
        String show_exact_match_default = props.getProperty("xcompare.displayexactmatch");
        if(show_exact_match_default != null && !show_exact_match_default.trim().isEmpty()) {
        	this.show_exact_match = Boolean.parseBoolean(show_exact_match_default);
        }
        
        String show_partial_exact_match_default = props.getProperty("xcompare.displaypartialexactmatch");
        if(show_partial_exact_match_default != null && !show_partial_exact_match_default.trim().isEmpty()) {
        	this.show_partial_exact_match = Boolean.parseBoolean(show_partial_exact_match_default);
        }
        
        String show_fuzzy_match_default = props.getProperty("xcompare.displayfuzzymatch");
        if(show_fuzzy_match_default != null && !show_fuzzy_match_default.trim().isEmpty()) {
        	this.show_fuzzy_match = Boolean.parseBoolean(show_fuzzy_match_default);
        }
        
        String show_no_match_default = props.getProperty("xcompare.displaynomatch");
        if(show_no_match_default != null && !show_no_match_default.trim().isEmpty()) {
        	this.show_no_match = Boolean.parseBoolean(show_no_match_default);
        }
        
        String show_extra_fuzzy_match_when_exact_match_found_default = props.getProperty("xcompare.displayextrafuzzymatchwhenexactmatchfound");
        if(show_extra_fuzzy_match_when_exact_match_found_default != null && !show_extra_fuzzy_match_when_exact_match_found_default.trim().isEmpty()) {
        	this.show_extra_fuzzy_match_when_exact_match_found = Boolean.parseBoolean(show_extra_fuzzy_match_when_exact_match_found_default);
        }
        
        String match_merged_reference_segments_default = props.getProperty("xcompare.matchmergedreferencesegments");
        if(match_merged_reference_segments_default != null && !match_merged_reference_segments_default.trim().isEmpty()) {
        	this.match_merged_reference_segments = Boolean.parseBoolean(match_merged_reference_segments_default);
        }
        
        String force_language_match_default = props.getProperty("xcompare.forcelanguagecodematch");
        if(force_language_match_default != null && !force_language_match_default.trim().isEmpty()) {
        	this.force_language_match = Boolean.parseBoolean(force_language_match_default);
        }
        
        String autoalign_previous_translation_default = props.getProperty("xcompare.autoalignprevioustranslation");
        if(autoalign_previous_translation_default != null && !autoalign_previous_translation_default.trim().isEmpty()) {
        	this.autoalign_previous_translation = Boolean.parseBoolean(autoalign_previous_translation_default);
        }
        
        String threshold_normal_default = props.getProperty("xcompare.threshold.noraml");
        if(threshold_normal_default != null && !threshold_normal_default.trim().isEmpty()) {
        	this.threshold_noraml = Integer.parseInt(threshold_normal_default);
        }
        
        String threshold_partialmatch_default = props.getProperty("xcompare.threshold.partialmatch");
        if(threshold_partialmatch_default != null && !threshold_partialmatch_default.trim().isEmpty()) {
        	this.threshold_partialmatch = Integer.parseInt(threshold_partialmatch_default);
        }
        
        in.close();
	}
	
	public HashMap<String, String> getReplacements(String sourceLanguageCode) throws IOException{
		HashMap<String, String> map = new HashMap<String, String>();
		
		FileInputStream in = new FileInputStream(this.configurationfile);
        Properties props = new Properties();
        props.load(in);
        
        String replacements_str = props.getProperty("xcompare.normalization.replace."+sourceLanguageCode);
        if(replacements_str == null) return map;
        
        String[] replacements = replacements_str.split(",");
        for(String replacement : replacements){
        	String[] pair = replacement.split("=");
        	if(pair.length == 2){
        		map.put(pair[0], pair[1]);
        	}
        }
        
        in.close();
		
		return map;
	}
}
