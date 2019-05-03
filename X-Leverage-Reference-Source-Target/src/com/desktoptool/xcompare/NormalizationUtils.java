package com.desktoptool.xcompare;

import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;

public class NormalizationUtils {

	private static List<String> weights = Arrays.asList("thousand","million","billion","trillion");
	private static List<String> hundreds = Arrays.asList("hundred");
	private static List<String> tens = Arrays.asList("twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety");
	private static List<String> ones = Arrays.asList("one","two","three","four","five","six","seven","eight","nine","ten"
			,"eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen");
	private static List<String> connectors = Arrays.asList("and");
	
	
	public static String normalize(String s, Configuration config, String sourceLanguageCode) {

		if (config.isNormalize_whitespace()){
			s = s.replaceAll("(\\s)+", " ");
			s = s.trim();
	    }
		
	    if (config.isNormalize_case()) {
	    	s = s.toLowerCase();
	    }
	    
	    if (config.isNormalize_englishtonumber()) {
	    	s = NormalizeEnglishWordsToNumbers(s);
	    }
	    
	    try{
	    	HashMap<String, String> replacements = config.getReplacements(sourceLanguageCode);
	    	for (String key : replacements.keySet()) {
	    		s = s.replace(key, (CharSequence)replacements.get(key));
	    	}
	    } catch (IOException e) {
	    	e.printStackTrace();
	    }
	    
	    if (config.isNormalize_punctuation())
	    {
	    	String[] nonlitrals = { ",", ".", "!", "?", "~", "'", "\"", ":", ";" };
	    	for (String punc : nonlitrals) {
	    		s = s.replace(punc, "");
	    	}
	    }
	    
	    return s;
	}
	
	public static String NormalizeEnglishWordsToNumbers(String s){
		
		StringBuffer sb = new StringBuffer();
		long sub_result = 0;
		long result = 0;
		String prev_connector = null;
		
		int level = -1;
		int sub_level = -1;
		int i = 0;
		while(i < s.length()) {
			
			if(s.charAt(i) == ' ' || s.charAt(i) == '-') {
				if(result == 0 && sub_result == 0) sb.append(s.charAt(i));
				i++;
			}else if(s.charAt(i) == ',' || s.charAt(i) == '.'){
				if(sub_result != 0 || result != 0){
					sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result));
					result = 0;
					sub_result = 0;
					
					sb.append(s.charAt(i));
					sub_level = -1;
				}
				i++;
			}else {
				String word = getNextWord(i,s).toLowerCase();
				i += word.length();
				if(word != null) {
					int cur_sub_level;
					int cur_level;
					if(ones.contains(word)) {
						cur_sub_level = 2;
						if(cur_sub_level > sub_level) {
							prev_connector = null;
						}else {
							if(sub_result != 0 || result != 0) {
								sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result) + " ");
								result = 0;
								sub_result = 0;
							}
						}
						sub_result += ones.indexOf(word)+1;
						sub_level = cur_sub_level;
						
					}else if(tens.contains(word)) {
						cur_sub_level = 1;
						if(cur_sub_level > sub_level) {
							prev_connector = null;
						}else {
							if(sub_result != 0 || result != 0) {
								sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result) + " ");
								result = 0;
								sub_result = 0;
							}
						}
						sub_result += (tens.indexOf(word)+2)*10;
						sub_level = cur_sub_level;
						
					}else if(hundreds.contains(word)) {
						cur_sub_level = 0;
						if(sub_level != -1) {
							sub_result *= 100;
							sub_level = cur_sub_level;
							prev_connector = null;
						}else {
							if(sub_result != 0 || result != 0) {
								sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result) + " ");
								result = 0;
								sub_result = 0;
							}
							sb.append(word);
							sub_level = -1;
						}
					}else if(weights.contains(word)) {
						cur_level = weights.size() - 1 - weights.indexOf(word);
						if(sub_level != -1 && cur_level > level) {
							result += sub_result*Math.pow(1000, weights.indexOf(word)+1);
							sub_result = 0;
							level = cur_level;
							prev_connector = null;
						}else {
							if(sub_result != 0 || result != 0) {
								sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result) + " ");
								result = 0;
								sub_result = 0;
							}
							sb.append(word);
							level = -1;
						}
						sub_level = -1;
					}else {
						if(connectors.contains(word)){
							if(sub_result != 0 || result != 0) {
								prev_connector = word;
							}else{
								sb.append(word);
								level = -1;
								sub_level = -1;
							}
						}else{
							if(sub_result != 0 || result != 0) {
								sb.append((prev_connector==null?"":prev_connector+" ") + (result+sub_result) + " ");
								result = 0;
								sub_result = 0;
							}
							sb.append(word);
							level = -1;
							sub_level = -1;
						}
					}
				}
			}
		}
		
		return sb.toString();
	}


	private static String getNextWord(int idx, String s) {
		StringBuffer sb = new StringBuffer();
		while(idx < s.length() && s.charAt(idx) != ' ' && s.charAt(idx) != ',' && s.charAt(idx) != '.' && s.charAt(idx) != '-') {
			sb.append(s.charAt(idx));
			idx++;
		}
		return sb.toString();
	}
}
