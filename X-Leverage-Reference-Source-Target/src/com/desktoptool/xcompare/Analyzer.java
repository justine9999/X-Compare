package com.desktoptool.xcompare;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.configuration.BaseConfiguration;
import org.apache.commons.configuration.Configuration;
import org.dom4j.DocumentHelper;
import org.dom4j.io.SAXReader;
import org.gs4tr.foundation.locale.Locale;
import org.gs4tr.tm.analysispackage.cleanup.cleaner.TxmlCleaner;
import org.gs4tr.tm.leverage.LeverageSupportImpl;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellArea;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Comment;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.OperatorType;
import com.aspose.cells.ProtectionType;
import com.aspose.cells.Range;
import com.aspose.cells.ShiftType;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import com.desktoptool.xcompare.LevenshteinDistance;
import com.desktoptool.xcompare.diff_match_patch;


public class Analyzer {

	private com.desktoptool.xcompare.Configuration config;
	private LeverageSupportImpl leverageSupport;
	private final static double Partial_match_score = 99.9D;
	private String tempfolder;
	private String populatefolder;
	private int jobid;
	private List<String> progresses;
	private StringBuffer sb;
	private int sourcefilecount;
	boolean isRunAgainstTargetReference = false;
	private String targetReport;
	
	private static String NULL_SOURCE_REFERENCE_STR = "X_COMPARE_NULL_SOURCE_REFERENCE_STR";
	//int count = 0;
	
	public String getTargetReport(){
		return this.targetReport;
	}
	
	public Analyzer(com.desktoptool.xcompare.Configuration config, int jobid, List<String> progresses, StringBuffer sb) {
		
		this.config = config;
		this.jobid = jobid;
		this.progresses = progresses;
		this.sb = sb;
	}
	
	public void setTempFolder(String tempfolder){
		this.tempfolder = tempfolder;
	}
	
	public void analyze(List<String> sourcefiles, List<String> sourceReferencefiles, List<String> targetReferencefiles, String languagePairCode, String outputDirectory) throws Exception {

		if(targetReferencefiles.size() != 0 && !config.isAutoalign_previous_translation()){
			this.isRunAgainstTargetReference = true;
		}
		
		String sourcelanguage = languagePairCode.split("_")[0];
		String targetlanguage = languagePairCode.split("_")[1];
		this.leverageSupport = new LeverageSupportImpl(Locale.makeLocale(sourcelanguage), Locale.makeLocale(targetlanguage));
		Workbook workbook = new Workbook();
		Worksheet infoTab = workbook.getWorksheets().get(0);
		Cells cells = infoTab.getCells();
		infoTab.setName("INDEX");
		infoTab.getCells().getColumns().get(0).setWidth(20.0D);
		infoTab.getCells().getColumns().get(1).setWidth(100.0D);
		
		List<XSegment> source_reference_xsegments = new ArrayList<XSegment>();
		for(String referencefile : sourceReferencefiles) {
			FileUtils.readSegments(referencefile, source_reference_xsegments, true, config);
		}
		
		List<XSegment> target_reference_xsegments = new ArrayList<XSegment>();
		for(String referencefile : targetReferencefiles) {
			FileUtils.readSegments(referencefile, target_reference_xsegments, true, config);
		}
		
		//align source reference and target reference text
		LinkedHashMap<String, String[]> align_results = new LinkedHashMap<String, String[]>();
		if(config.isAutoalign_previous_translation()){
			sb.append("[aligning source reference files and target reference files]\n");
			progresses.set(this.jobid, "A");
			align_results = AutoAlignSourceAndTargetReference(source_reference_xsegments, target_reference_xsegments, sourcelanguage, targetlanguage);
			progresses.set(this.jobid, "0.0%");
		}
		
		//clean populated folder
		this.populatefolder = outputDirectory + File.separator + "X-Compare_populate_" + FileUtils.getNameWithoutExtension(sourcefiles.get(0)) + "_" + languagePairCode.toUpperCase() + "_[" + System.currentTimeMillis()  + "]";
		if(!new File(this.populatefolder).exists()){
			new File(this.populatefolder).mkdirs();
		}else{
			org.apache.commons.io.FileUtils.cleanDirectory(new File(this.populatefolder));
		}
	
		this.sourcefilecount = sourcefiles.size();
		int index = 0;
		for(String sourcefile : sourcefiles) {
			
			String tabName = "Tab - " + index;
			infoTab.getCells().get(index, 0).setValue(tabName);
			infoTab.getHyperlinks().add(index, 0, 1, 1, "'" + tabName+"'!A1");
			infoTab.getCells().get(index, 1).setValue(new File(sourcefile).getName());
			cells.setRowHeight(index, 30.0D);
			
			Worksheet sheet = workbook.getWorksheets().add(tabName);
			sheet.setZoom(75);
			
			String orphanTabName = "Tab - " + index + " - unmatched";
			Worksheet sheet_orphan = workbook.getWorksheets().add(orphanTabName);
			
			analyze_internal(sourcefile, source_reference_xsegments, target_reference_xsegments, sheet, sheet_orphan, sourcelanguage, targetlanguage, align_results, index);
			
			sheet.protect(ProtectionType.ALL, null, null);
			index++;
		}
		
		//index tab tab name column format
		Range range = cells.createRange(0,0,index,1);
		range.setColumnWidth(15.0D);
		Style style = workbook.createStyle();
		style.getFont().setColor(com.aspose.cells.Color.getBlue());
		style.getFont().setUnderline(1);
		style.setTextWrapped(true);
		style.setHorizontalAlignment(TextAlignmentType.CENTER);
		style.setVerticalAlignment(TextAlignmentType.CENTER);
		style.setPattern(1);
		style.setForegroundColor(com.aspose.cells.Color.fromArgb(230, 230, 230));
		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		StyleFlag styleFlag = new StyleFlag();
		styleFlag.setAll(true);
		range.applyStyle(style, styleFlag);
		
		//index tab file name column format
		range = cells.createRange(0,1,index,1);
		range.setColumnWidth(100.0D);
		style = workbook.createStyle();
		style.setTextWrapped(true);
		style.setHorizontalAlignment(TextAlignmentType.LEFT);
		style.setIndentLevel(1);
		style.setVerticalAlignment(TextAlignmentType.CENTER);
		style.setPattern(1);
		style.setForegroundColor(com.aspose.cells.Color.fromArgb(230, 230, 230));
		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		styleFlag = new StyleFlag();
		styleFlag.setAll(true);
		range.applyStyle(style, styleFlag);
		
		String finalreport = outputDirectory + File.separator + "X-Compare_" + FileUtils.getNameWithoutExtension(sourcefiles.get(0)) + "_" + languagePairCode.toUpperCase() + "_[" + System.currentTimeMillis()  + "]" + ".xlsx";
		workbook.save(finalreport);
		this.targetReport = finalreport;
		
		//if run against old excel report
		if(source_reference_xsegments.size() == 0){
			String previousrerpot = sourceReferencefiles.get(0);
			String populatedfile = new File(this.populatefolder).listFiles()[0].getAbsolutePath();
			populateOriginalReportToNewReportAndSourceFile(previousrerpot, finalreport, populatedfile, sourcelanguage, targetlanguage);
		}
		
		//create TM
		/*for(File populatedfile : new File(populatedfolder).listFiles()){
			
			String ext = FileUtils.getExtension(populatedfile.getName());
			if(ext.toLowerCase().equals("txml")){
				createTM(populatedfile.getAbsolutePath(), sourcelanguage, targetlanguage);
			}	
		}*/
		//System.out.println("count: " + this.count);
	}
	
	private void populateOriginalReportToNewReportAndSourceFile(String orgreport, String newreport, String populatedfile, String sourcelanguage, String targetlanguage) throws Exception{
		
		Workbook ow = new Workbook(orgreport);
		Workbook cw = new Workbook(newreport);
		cw.getWorksheets().get(1).getComments().clear();
		ow.getWorksheets().get(1).getComments().clear();

		Cells ocells = ow.getWorksheets().get(1).getCells();
		Cells ccells = cw.getWorksheets().get(1).getCells();
		
		ccells.clearContents(0, 0, ccells.getMaxRow(), 6);
		ccells.clearFormats(0, 0, ccells.getMaxRow(), 6);
		ccells.clearRange(0, 0, ccells.getMaxRow(), 6);
		for(Object ocellarea : cw.getWorksheets().get(1).getCells().getMergedCells()){
			CellArea cellarea = (CellArea)ocellarea;
			ccells.get(cellarea.StartRow, cellarea.StartColumn).getMergedRange().unMerge();
		}
		
		for(int c = 0; c <= 6; c++){	
			for(int r = 0; r <= ocells.getMaxRow(); r++){
				Cell ccell = ccells.get(r, c);
				Cell ocell = ocells.get(r, c);
				if(c == 3){
					ccell.setHtmlString(ocell.getHtmlString());
				}else{
					ccell.setValue(ocell.getValue());
				}
				ccell.setStyle(ocell.getStyle());
			}
		}
		
		
		if(config.isAutoalign_previous_translation()){
			
			Cells ocells_orphan = ow.getWorksheets().get(2).getCells();
			Cells ccells_orphan = cw.getWorksheets().get(2).getCells();
			ccells_orphan.clearContents(0, 0, ccells_orphan.getMaxRow(), ccells_orphan.getMaxColumn());
			ccells_orphan.clearFormats(0, 0, ccells_orphan.getMaxRow(), ccells_orphan.getMaxColumn());
			ccells_orphan.clearRange(0, 0, ccells_orphan.getMaxRow(), ccells_orphan.getMaxColumn());
			for(int c = 0; c <= 3; c++){
				for(int r = 0; r <= ocells_orphan.getMaxRow(); r++){
					Cell ccell_orphan = ccells_orphan.get(r, c);
					Cell ocell_orphan = ocells_orphan.get(r, c);
					if(c <= 1){
						ccell_orphan.setValue(ocell_orphan.getValue());
					}else{
						ccell_orphan.setFormula(ocell_orphan.getFormula());
					}
					
					ccell_orphan.setStyle(ocell_orphan.getStyle());
				}
			}
		}
		
		if(config.isAutoalign_previous_translation()){
			
			for(int c = 7; c <= 12; c++){	
				for(int r = 0; r <= ocells.getMaxRow(); r++){
					Cell ccell = ccells.get(r, c);
					Cell ocell = ocells.get(r, c);
					if(r > 0 && c == 7){
						ccell.setFormula(ocell.getFormula());
					}else{
						ccell.setValue(ocell.getValue());
					}
					ccell.setStyle(ocell.getStyle());
				}
			}
			
			cw.calculateFormula();
		}
		
		//source col 1, target col 7
		int start = 1;
		HashMap<String, String> notes = new HashMap<String, String>();
		HashMap<String, String> conclusions = new HashMap<String, String>();
		HashMap<String, String> alignedtargets = new HashMap<String, String>();
		HashMap<String, String> scores = new HashMap<String, String>();
		
		while(true){
			
			String segid = ccells.get(start, 0).getStringValue();
			
			//shift non-empty matches up
			shiftRange(ccells, start, 1);
			shiftRange(ccells, start, 7);
			
			// for alignment
			int oidx = getEndOfSource(start, ccells, 1);
			int cidx = getEndOfSource(start, ccells, 7);
			
			recreateTracksAndScores(start, oidx, ccells, 1, sourcelanguage);
			recreateTracksAndScores(start, cidx, ccells, 7, targetlanguage);
			
			if(config.isAutoalign_previous_translation()){
				String targetstringval = ccells.get(start, 7).getDisplayStringValue();
				alignedtargets.put(segid, targetstringval.equals("#N/A")?"":targetstringval.substring(0, targetstringval.lastIndexOf('-')).trim());
				String alignedscore_str = targetstringval.equals("#N/A")?"[ 0 ]":targetstringval.substring(targetstringval.lastIndexOf('-')+1, targetstringval.length()).trim();
				int alignedscore = Integer.parseInt(alignedscore_str.substring(1, alignedscore_str.length()-1).trim());
				String score_text = ccells.get(start, 4).getDisplayStringValue().trim();
				int matchscore = score_text.equals("100-")?75:Integer.parseInt(score_text);
				matchscore = matchscore == 100?99:80;
				int finalscore = alignedscore == 0?0:matchscore;
				scores.put(segid, Integer.toString(finalscore));
			}
						
			deleteEmptyRange(oidx+1, ccells, 1, 0, 6);
			boolean hasTarget = (ccells.get(start, 7).getDisplayStringValue() != null && !ccells.get(start, 7).getDisplayStringValue().trim().equals("")) || (ccells.get(start, 7).getFormula() != null && !ccells.get(start, 7).getFormula().trim().equals(""));
			if(hasTarget){
				deleteEmptyRange(cidx+1, ccells, 7, 7, 12);
			}

			if(oidx > cidx){
				insertEmptyRange(ccells, cidx+1, oidx, 7, 12);
			}else if(cidx > oidx){
				insertEmptyRange(ccells, oidx+1, cidx, 0, 6);
			}
			
			int max_span = Math.max(oidx, cidx) - start + 1;
			Range range_id = ccells.createRange(start, 0, max_span, 1);
			range_id.merge();
			Range range_source = ccells.createRange(start, 1, max_span, 1);
			range_source.merge();
			Range range_source_notes = ccells.createRange(start, 5, max_span, 1);
			range_source_notes.merge();
			Range range_target = ccells.createRange(start, 7, max_span, 1);
			range_target.merge();
			Range range_target_notes = ccells.createRange(start, 11, max_span, 1);
			range_target_notes.merge();
			
			int final_source_match_cnt = countMathces(start, start+max_span-1, 6, ccells);
			int final_target_match_cnt = countMathces(start, start+max_span-1, 12, ccells);
			formatSourceIdCell(start, max_span, final_source_match_cnt, final_target_match_cnt, ccells, hasTarget);
			
			
			// for populating source file
			String score = ccells.get(start, 4).getStringValue().trim();
			String note = ccells.get(start, 5).getStringValue().trim();
			if(!note.equals("")){
				notes.put(segid, note);
			}else{
				conclusions.put(segid, score);
			}
			
			start += max_span;
			
			if(ccells.get(start, 1).getStringValue().equals("") 
					&& ccells.get(start, 2).getStringValue().equals("")
						&& ccells.get(start, 7).getStringValue().equals("")
							&& ccells.get(start, 8).getStringValue().equals("")){
				break;
			}
		}
		
		if(config.isAutoalign_previous_translation()){
			cw.getWorksheets().get(2).getAutoFilter().setRange("A1:D1");
		}
		
		cw.save(newreport);
		
		FileUtils.populateReportNotesToNotesAndAlignedTargets(populatedfile, notes, conclusions, alignedtargets, scores, config);
	}
	
	private int getEndOfSource(int start, Cells cells, int col){
		
		while ((cells.get(start + 1, col).getStringValue().trim().equals("")) && 
			      (!cells.get(start + 1, col + 1).getStringValue().trim().equals(""))) {
			      start++;
			}
		return start;
	}
	
	private void shiftRange(Cells cells, int startrow, int src_col) throws Exception{
		
		int endrow = startrow;
		while(cells.get(endrow + 1, src_col).getStringValue().trim().equals("")
				&& !cells.get(endrow + 1, src_col+3).getStringValue().trim().equals("")){
			endrow++;
		}
		int head = startrow;
		int i = startrow;
		while(i <= endrow){
			if(!cells.get(i, src_col+1).getStringValue().trim().equals("")){
				
				cells.get(head, src_col+1).setValue(cells.get(i, src_col+1).getStringValue());
				cells.get(head, src_col+2).setHtmlString(cells.get(i, src_col+2).getHtmlString());
				cells.get(head, src_col+3).setValue(cells.get(i, src_col+3).getStringValue());
				cells.get(head, src_col+5).setValue(cells.get(i, src_col+5).getStringValue());
				
				if(head != i){
					cells.get(i, src_col+1).setValue(null);
					cells.get(i, src_col+2).setValue(null);
					cells.get(i, src_col+5).setValue(null);
				}
				
				head++;
			}
			i++;
		}
	}
	
	private void deleteEmptyRange(int idx, Cells cells, int checkcol, int startcol, int endcol){
		
	    while(idx < cells.getMaxRow() && (cells.get(idx, checkcol).getStringValue().trim().equals("")) && 
	    		(cells.get(idx, checkcol).getFormula() == null || cells.get(idx, checkcol).getFormula().trim().equals("")) &&
	    			(cells.get(idx, checkcol + 1).getStringValue().trim().equals(""))){
	    	
	      cells.deleteRange(idx, startcol, idx, endcol, ShiftType.UP);
	    }
	}

	private void insertEmptyRange(Cells cells, int startrow, int endrow, int startcol, int endcol){
		//System.out.println("test: " + startrow + " " + endrow + " " + startcol + " " + endcol);
		cells.insertRange(CellArea.createCellArea(startrow, startcol, endrow, endcol), ShiftType.DOWN);
		for(int i = startrow; i <= endrow; i++){
			Cell score_cell = cells.get(i, endcol-2);
			score_cell.setValue(0);
			formatSimilarityScoreCell(score_cell, cells);
		}
	}
	
	private int countMathces(int start, int end, int col, Cells cells){
		
		int cnt = 0;
		
		StringBuffer sb = new StringBuffer();
		for(int i = start; i <= end; i++){
			//match created by x-compare
			int cnt_auto_add = 0;
			Cell cell = cells.get(i, col);
			String text = cell.getDisplayStringValue();
			boolean isIdxArea = false;
			for(int j = 0; j < text.length(); j++){
				char c = text.charAt(j);
				if(c == '<'){
					isIdxArea = true;
				}else if(c == '>'){
					isIdxArea = false;
					cnt_auto_add += sb.toString().split("\\,").length;
					sb = new StringBuffer();
				}else{
					if(isIdxArea){
						sb.append(c);
					}
				}
			}
			
			if(cnt_auto_add == 0){
				if(!cells.get(i, col-4).getDisplayStringValue().trim().equals("")) {
					cnt++;
				}
			}else{
				cnt += cnt_auto_add;
			}	
		}
		
		return cnt;
	}
	
	private void recreateTracksAndScores(int start, int end, Cells cells, int col, String language) {
		
		String cell_src_str = cells.get(start, col).getDisplayStringValue().trim();
		for(int i = start; i <= end; i++){

			String cell_ref_str = cells.get(i, col+1).getDisplayStringValue().trim();
			
			if(!cell_ref_str.equals("")){
				
				String cell_src_str_normalize = NormalizationUtils.normalize(cell_src_str, config, language);
				String cell_ref_str_normalize = NormalizationUtils.normalize(cell_ref_str, config, language);

				double[] scores = getMatchScore(cell_src_str_normalize, cell_ref_str_normalize, null, null, null, null, language);
				
				String trackedString = "";
				if(scores[1] == Partial_match_score){
					//trackedString = createHtmlForPartialExactMatch(source_xsegment, xsegment, 0);
					trackedString = createHtmlForFuzzyMatch(cell_ref_str, cell_src_str);
					cells.get(i, col+3).setValue("100-");
				}else{
					trackedString = createHtmlForFuzzyMatch(cell_ref_str, cell_src_str);
					cells.get(i, col+3).setValue((int)scores[0]);
				}
				
				cells.get(i, col+2).setHtmlString(trackedString);

			} else {
				cells.get(i, col+3).setValue(0);
				cells.get(i, col+2).setHtmlString("");
			}
			
			formatReferenceCell(cells.get(i, col+1), cells);
			formatSimilarityScoreCell(cells.get(i, col+3), cells);
		}
	}
	
	private void analyze_internal(String sourcefile, List<XSegment> source_reference_xsegments, List<XSegment> target_reference_xsegments, Worksheet sheet, Worksheet sheet_orphan, String sourcelanguage, String targetlanguage, LinkedHashMap<String, String[]> align_results, int index) throws Exception {
		
		List<XSegment> source_xsegments = new ArrayList<XSegment>();
		List<XSegment> target_xsegments = new ArrayList<XSegment>();
		FileUtils.readSegments(sourcefile, source_xsegments, true, config);
		FileUtils.readSegments(sourcefile, target_xsegments, false, config);
		
		Cells cells = sheet.getCells();
		cells.get(0, 0).setValue("Source Segment ID");
		cells.get(0, 1).setValue("Source Text");
		cells.get(0, 2).setValue("Source Reference Text");
		cells.get(0, 3).setValue("Tracked Change(s)");
		cells.get(0, 4).setValue("Similarity Score");
		cells.get(0, 5).setValue("Notes");
		cells.get(0, 6).setValue("Source Reference File Name & Segment ID");
		
		cells.get(0, 7).setValue("Target Text");
		cells.get(0, 8).setValue("Target Reference Text");
		cells.get(0, 9).setValue("Tracked Change(s)");
		cells.get(0, 10).setValue("Similarity Score");
		cells.get(0, 11).setValue("Notes");
		cells.get(0, 12).setValue("Target Reference File Name & Segment ID");
		
		int cell_start_index = 1;
		
		//for tm creation 0-no match, 1-fuzzy match, 2-partial exact match 3-exact match
		int[] match_conclusions = new int[source_xsegments.size()];
		
		String[] aligned_targets = new String[source_xsegments.size()];
		int[] aligned_final_score = new int[source_xsegments.size()];
		
		//store the boundaries of the concatenated string
		int[][] src_boundaries = new int[source_reference_xsegments.size()][2];
		String src_concatReferenceStr = concatenateAllReferenceText(source_reference_xsegments, src_boundaries, false);
				
		//store the boundaries of the concatenated normalized string
		int[][] src_boundaries_normalized = new int[source_reference_xsegments.size()][2];
		String src_concatReferenceStr_normalized = concatenateAllReferenceText(source_reference_xsegments, src_boundaries_normalized, true);
		
		//store the boundaries of the concatenated string
		int[][] trg_boundaries = new int[target_reference_xsegments.size()][2];
		String trg_concatReferenceStr = concatenateAllReferenceText(target_reference_xsegments, trg_boundaries, false);
				
		//store the boundaries of the concatenated normalized string
		int[][] trg_boundaries_normalized = new int[target_reference_xsegments.size()][2];
		String trg_concatReferenceStr_normalized = concatenateAllReferenceText(target_reference_xsegments, trg_boundaries_normalized, true);
		
		//store used aligned entry, key is the source reference text
		HashSet<String> usedAlignedEntrySource = new HashSet<String>();
		
		int end_cell_index = 0;

		for(int i = 0; i < source_xsegments.size(); i++) {
			
			String p = String.format("%.1f%%", (((i+1)*100*1.0/source_xsegments.size())*((index+1)/this.sourcefilecount)));
			this.progresses.set(this.jobid, p);
			
			//gather source and source reference segments for comparison
			XSegment source_xsegment = source_xsegments.get(i);
			List<XSegment> source_similar_reference_segments = new ArrayList<XSegment>();
			List<String> source_diff_texts = new ArrayList<String>();
			List<double[]> source_scores = new ArrayList<double[]>();
			HashSet<Integer> idx_exact_match_source = new HashSet<Integer>();
			
			findReferenceMatches(source_xsegment, source_reference_xsegments, src_concatReferenceStr, src_boundaries, src_concatReferenceStr_normalized, src_boundaries_normalized, match_conclusions, i, source_similar_reference_segments, source_diff_texts, source_scores, idx_exact_match_source, sourcelanguage);


			//gather target and target reference segments for comparison
			XSegment target_xsegment = target_xsegments.get(i);
			List<XSegment> target_similar_reference_segments = new ArrayList<XSegment>();
			List<String> target_diff_texts = new ArrayList<String>();
			List<double[]> target_scores = new ArrayList<double[]>();
			HashSet<Integer> idx_exact_match_target = new HashSet<Integer>();

			findReferenceMatches(target_xsegment, target_reference_xsegments, trg_concatReferenceStr, trg_boundaries, trg_concatReferenceStr_normalized, trg_boundaries_normalized, match_conclusions, i, target_similar_reference_segments, target_diff_texts, target_scores, idx_exact_match_target, targetlanguage);
			
			//get rid of extra fuzzies
			if(!config.isShow_extra_fuzzy_match_when_exact_match_found()) {
				if(idx_exact_match_source.size() != 0){
					removeAllOtherElementExcept(source_similar_reference_segments, idx_exact_match_source);
					removeAllOtherElementExcept(source_diff_texts, idx_exact_match_source);
					removeAllOtherElementExcept(source_scores, idx_exact_match_source);
				}
				
				if(idx_exact_match_target.size() != 0){
					removeAllOtherElementExcept(target_similar_reference_segments, idx_exact_match_target);
					removeAllOtherElementExcept(target_diff_texts, idx_exact_match_target);
					removeAllOtherElementExcept(target_scores, idx_exact_match_target);
				}
			}
			
			//consolidate reference semgents and their infos
			LinkedHashMap<String, HashMap<String, List<String>>> src_ref_map = consolidateReferenceFilesInfo(source_similar_reference_segments, source_diff_texts, source_scores);
			LinkedHashMap<String, HashMap<String, List<String>>> trg_ref_map = consolidateReferenceFilesInfo(target_similar_reference_segments, target_diff_texts, target_scores);
			
			int final_src_match_cnt = source_similar_reference_segments.size();
			int final_trg_match_cnt = target_similar_reference_segments.size();
			
			int max_span = Math.max(1, Math.max(src_ref_map.size(), trg_ref_map.size()));
			
			boolean hasTarget = !target_xsegment.getContent().trim().equals("");
			formatSourceIdCell(cell_start_index, max_span, final_src_match_cnt, final_trg_match_cnt, cells, hasTarget);
			
			int org_cell_start_index = cell_start_index;
			//write source comparison into excel
			String aligned_trg_text = "";
			int align_score = 0;
			if(src_ref_map.size() > 0 && isComparable(source_xsegment.getNormalized_content())){
				
				cells.get(cell_start_index, 0).setValue(Integer.parseInt(source_xsegment.getSegment_id()));
				cells.get(cell_start_index, 1).setValue(source_xsegment.getContent());
				
				String used_ref_text = null;
				
				for(String text : src_ref_map.keySet()){
					String ref_text = text.split("#%#")[0];
					String diff_text = text.split("#%#")[1];
					String score_text = text.split("#%#")[2];
					String ref_normalized_text = text.split("#%#")[3];
					cells.get(cell_start_index, 2).setValue(ref_text);
					cells.get(cell_start_index, 3).setHtmlString(diff_text);
					cells.get(cell_start_index, 4).setValue(score_text.equals("100-")?score_text:Integer.parseInt(score_text));
					
					formatSimilarityScoreCell(cells.get(cell_start_index, 4), cells);
					
					HashMap<String, List<String>> metas = src_ref_map.get(text);
					StringBuffer sb = new StringBuffer();
					for(String filename : metas.keySet()){
						sb.append("● " + filename + " < ");
						List<String> ids = metas.get(filename);
						for(int z = 0; z < ids.size(); z++){
							sb.append(ids.get(z));
							if(z != ids.size()-1) sb.append(',');
						}
						sb.append(" >");
						sb.append("\n");
					}
					cells.get(cell_start_index, 6).setValue(sb.toString());

					cell_start_index++;
					
					//get highest match from aligned target reference text
					if(align_results.containsKey(ref_text)){
						int source_match_score = score_text.equals("100-")?75:Integer.parseInt(score_text);
						source_match_score = source_match_score == 100?99:80;
						int autoaligner_confidence_score = Integer.parseInt(align_results.get(ref_text)[1]);
						//int final_score = autoaligner_confidence_score * source_match_score / 100;
						int final_score = autoaligner_confidence_score == 0?0:source_match_score;
						if(final_score > align_score && final_score >= config.getThreshold_noraml()){
							align_score = final_score;
							used_ref_text = ref_text;
							aligned_trg_text = align_results.get(ref_text)[0];
						}
					}
				}
				
				if(used_ref_text != null){
					
					usedAlignedEntrySource.add(used_ref_text);
				}
				
				for(int k = src_ref_map.size(); k < max_span; k++) {
					cells.get(cell_start_index, 4).setValue(0);
					formatSimilarityScoreCell(cells.get(cell_start_index, 4), cells);
					cell_start_index++;
				}		
				
			}else{
				
				cells.get(cell_start_index, 0).setValue(Integer.parseInt(source_xsegment.getSegment_id()));
				cells.get(cell_start_index, 1).setValue(source_xsegment.getContent());

				for(int k = 0; k < max_span; k++) {
					
					cells.get(cell_start_index, 2).setValue("");
					cells.get(cell_start_index, 3).setValue("");
					cells.get(cell_start_index, 4).setValue(0);
					
					formatSimilarityScoreCell(cells.get(cell_start_index, 4), cells);
					
					cells.get(cell_start_index, 5).setValue("");
					cells.get(cell_start_index, 6).setValue("");
					
					cell_start_index++;
				}
			}
			
			Range range_source_index = cells.createRange(cell_start_index-max_span, 0, max_span, 1);
			range_source_index.merge();
			
			Range range_source_text = cells.createRange(cell_start_index-max_span, 1, max_span, 1);
			range_source_text.merge();
			
			Range range_source_notes = cells.createRange(cell_start_index-max_span, 5, max_span, 1);
			range_source_notes.merge();
			
			
			end_cell_index = cell_start_index;
			cell_start_index = org_cell_start_index;
			//write target comparison into excel
			
			if(config.isAutoalign_previous_translation()){
				
				cells.get(cell_start_index, 7).setFormula("=INDEX('" + sheet_orphan.getName() + "'!B2:B" + (align_results.size()+1) + ",MATCH(TRUE,INDEX('" + sheet_orphan.getName() + "'!A2:A" + (align_results.size()+1) + "=C" + (cell_start_index+1) + ",0),0))");
				aligned_targets[i] = aligned_trg_text;
				aligned_final_score[i] = align_score;
				
				for(int k = 0; k < max_span; k++) {
					
					cells.get(cell_start_index, 8).setValue("");
					cells.get(cell_start_index, 9).setValue("");
					cells.get(cell_start_index, 10).setValue(0);
					
					formatSimilarityScoreCell(cells.get(cell_start_index, 10), cells);
					
					cells.get(cell_start_index, 11).setValue("");
					cells.get(cell_start_index, 12).setValue("");
					
					cell_start_index++;
				}
			}else{
				
				if(trg_ref_map.size() > 0 && isComparable(target_xsegment.getNormalized_content())){
					
					cells.get(cell_start_index, 7).setValue(target_xsegment.getContent());
					
					for(String text : trg_ref_map.keySet()){
						String ref_text = text.split("#%#")[0];
						String diff_text = text.split("#%#")[1];
						String score_text = text.split("#%#")[2];
						cells.get(cell_start_index, 8).setValue(ref_text);
						cells.get(cell_start_index, 9).setHtmlString(diff_text);
						cells.get(cell_start_index, 10).setValue(score_text.equals("100-")?score_text:Integer.parseInt(score_text));
						
						formatSimilarityScoreCell(cells.get(cell_start_index, 10), cells);
						
						HashMap<String, List<String>> metas = trg_ref_map.get(text);
						StringBuffer sb = new StringBuffer();
						for(String filename : metas.keySet()){
							sb.append("● " + filename + " < ");
							List<String> ids = metas.get(filename);
							for(int z = 0; z < ids.size(); z++){
								sb.append(ids.get(z));
								if(z != ids.size()-1) sb.append(',');
							}
							sb.append(" >");
							sb.append("\n");
						}
						cells.get(cell_start_index, 12).setValue(sb.toString());
						cell_start_index++;
					}
					
					for(int k = trg_ref_map.size(); k < max_span; k++) {
						cells.get(cell_start_index, 10).setValue(0);
						formatSimilarityScoreCell(cells.get(cell_start_index, 10), cells);
						cell_start_index++;
					}		
					
				}else{
					
					cells.get(cell_start_index, 7).setValue(target_xsegment.getContent());

					for(int k = 0; k < max_span; k++) {
						
						cells.get(cell_start_index, 8).setValue("");
						cells.get(cell_start_index, 9).setValue("");
						cells.get(cell_start_index, 10).setValue(0);
						
						formatSimilarityScoreCell(cells.get(cell_start_index, 10), cells);
						
						cells.get(cell_start_index, 11).setValue("");
						cells.get(cell_start_index, 12).setValue("");
						
						cell_start_index++;
					}
				}
				
			}
			
			
			Range range_target_text = cells.createRange(cell_start_index-max_span, 7, max_span, 1);
			range_target_text.merge();
			
			Range range_target_notes = cells.createRange(cell_start_index-max_span, 11, max_span, 1);
			range_target_notes.merge();
		}	
		
		//System.out.println();
		
		formatExcelSheet(sheet, cell_start_index);
		
		//create txmls with populated notes
		String populatedfile = this.populatefolder + File.separator + new File(sourcefile).getName();
		org.apache.commons.io.FileUtils.copyFile(new File(sourcefile), new File(populatedfile));
		FileUtils.fixTxmlTxlfLanguageCodes(populatedfile, sourcelanguage, targetlanguage);
		FileUtils.populateMatchConclusionToNotes(populatedfile, match_conclusions);
		if(config.isAutoalign_previous_translation()){
			FileUtils.populateAlignedOldTargets(populatedfile, aligned_targets, aligned_final_score, config);
		}
		
		populateOrphanTab(sheet, end_cell_index-1, sheet_orphan, align_results, usedAlignedEntrySource);
	}
	
	private void populateOrphanTab(Worksheet sheet, int end_cell_index, Worksheet sheet_orphan, LinkedHashMap<String, String[]> align_results, HashSet<String> usedAlignedEntrySource){
		
		Cells cells = sheet_orphan.getCells();
		cells.get(0,0).setValue("Source Reference Text");
		cells.get(0,1).setValue("Target Text");
		cells.get(0,2).setValue("Source Reference Used");
		cells.get(0,3).setValue("Target Used");
		formatOrphanHeaderRow(sheet_orphan, cells);
		
		Iterator it = align_results.entrySet().iterator();
		int idx = 1;
		while(it.hasNext()){
			Map.Entry<String, String[]> entry = (Map.Entry<String, String[]>)it.next();
			String source = entry.getKey().replaceAll("\\[repeat\\]+$", "");
			String target = entry.getValue()[0];
			String score = entry.getValue()[1];
			
			if(!source.startsWith(NULL_SOURCE_REFERENCE_STR)){
				cells.get(idx, 0).setValue(source);
			}else{
				cells.get(idx, 0).setValue("--------");
			}
			
			//conditional formatting
			int index_s = sheet_orphan.getConditionalFormattings().add();
			FormatConditionCollection fcs_s = sheet_orphan.getConditionalFormattings().get(index_s);
			CellArea ca_s = new CellArea();
			ca_s.StartRow = ca_s.EndRow = idx;
			ca_s.StartColumn = ca_s.EndColumn = 0;
			fcs_s.addArea(ca_s);
			int fcindex_s = fcs_s.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "=IF(C" + (idx+1) + "=TRUE,TRUE,FALSE)", "");
			fcs_s.get(fcindex_s).getStyle().setBackgroundColor(com.aspose.cells.Color.fromArgb(255, 243, 203));
			
			cells.get(idx, 1).setValue(target + " - [ " + score + " ]");
			
			//conditional formatting
			int index_t = sheet_orphan.getConditionalFormattings().add();
			FormatConditionCollection fcs_t = sheet_orphan.getConditionalFormattings().get(index_t);
			CellArea ca_t = new CellArea();
			ca_t.StartRow = ca_t.EndRow = idx;
			ca_t.StartColumn = ca_t.EndColumn = 1;
			fcs_t.addArea(ca_t);
			int fcindex_t = fcs_t.addCondition(FormatConditionType.EXPRESSION, OperatorType.NONE, "=IF(D" + (idx+1) + "=TRUE,TRUE,FALSE)", "");
			fcs_t.get(fcindex_t).getStyle().setBackgroundColor(com.aspose.cells.Color.fromArgb(255, 243, 203));
			
			cells.get(idx, 2).setFormula("=NOT(ISERROR(MATCH(TRUE,INDEX('" + sheet.getName() + "'!C2:C" + end_cell_index + "=A" + (idx+1) + ",0),0)))");
			cells.get(idx, 3).setFormula("=NOT(ISERROR(MATCH(TRUE,INDEX('" + sheet.getName() + "'!H2:H" + end_cell_index + "=B" + (idx+1) + ",0),0)))");
			
			formatOrphanRow(sheet_orphan, cells, idx, usedAlignedEntrySource.contains(source));
			
			idx++;
		}
		
		sheet_orphan.freezePanes(1, 0, 1, 30);
	}
	
	private String concatenateAllReferenceText(List<XSegment> reference_xsegments, int[][] boundaries, boolean isNormalize) {
		
		StringBuffer sb = new StringBuffer();
		for(int i = 0; i < reference_xsegments.size(); i++){
			String str = isNormalize?reference_xsegments.get(i).getNormalized_content():reference_xsegments.get(i).getContent();
			if(i == 0){
				sb.append(str);
				boundaries[i] = new int[]{0, str.length()-1};
			}else{
				sb.append((isNormalize?" ":"\n") + str);
				boundaries[i] = new int[]{boundaries[i-1][1]+2, boundaries[i-1][1]+2+str.length()-1};
			}
		}
		
		return sb.toString();
	}
	
	private void findReferenceMatches(XSegment xsegment, List<XSegment> reference_xsegments, String refStr, int[][] boundaries, String refStrNormalized, int[][] boundariesNormalized, int[] match_conclusions, int i, List<XSegment> similar_reference_segments, List<String> diff_texts, List<double[]> scores, HashSet<Integer> idx_exact_match, String language){
		
		String text = xsegment.getContent();
		String normalized_text = xsegment.getNormalized_content();
		if(normalized_text.length() == 0) return;

		int span = config.isMatch_merged_reference_segments()?reference_xsegments.size():1;
		//combine match to solve segmentation issue for source reference file
		//System.out.println("ttttt:  "+normalized_source_text);
		int end = -1;
		for(int j = 0; j < reference_xsegments.size(); j++) {
			double[] prev_scores = new double[2];
			
			int k = j;
			while(k < reference_xsegments.size() && k < j+span){
				
				double[] cur_scores = getMatchScore(refStrNormalized, normalized_text, boundariesNormalized[j][0], boundariesNormalized[k][1], null, null, language);
				//System.out.println("ppppp:  " + cur_normalize + "      " + normalized_source_text);
				//System.out.println("mmmmmm:  " + cur_scores[0] + "      " + cur_normalize);
				if(normalized_text.length()*100/(boundariesNormalized[k][1]-boundariesNormalized[j][0]+1) < config.getThreshold_noraml() || cur_scores[0] <= prev_scores[0]) {
					break;
				}
				
				prev_scores = cur_scores;

				k++;
			}

			if(prev_scores[0] >= config.getThreshold_noraml()){

				if(j > end || scores.size() == 0 || scores.get(scores.size()-1)[0] < prev_scores[0]){
					
					if(j <= end && scores.get(scores.size()-1)[0] < prev_scores[0]){
						similar_reference_segments.remove(similar_reference_segments.size()-1);
						scores.remove(scores.size()-1);
						diff_texts.remove(diff_texts.size()-1);
					}
					
					match_conclusions[i] = Math.max(prev_scores[0]<Partial_match_score?1:(prev_scores[1]==Partial_match_score?2:3), match_conclusions[i]);
					
					if(!config.isShow_exact_match() && prev_scores[0] == 100){
						continue;
					}
					
					if(!config.isShow_fuzzy_match() && prev_scores[0] < 100 && prev_scores[1] != Partial_match_score){
						continue;
					}
					
					if(!config.isShow_partial_exact_match() && prev_scores[1] == Partial_match_score){
						continue;
					}
					
					if(prev_scores[0] == 100){
						idx_exact_match.add(scores.size());
					}
					
					int s = Integer.parseInt(reference_xsegments.get(j).getSegment_id());
					String seg_range = (k-j-1==0)?(""+s):(s + " - " + (s+k-j-1));
					String str = refStr.substring(boundaries[j][0], boundaries[k-1][1]+1);
					String strNormalized = refStrNormalized.substring(boundariesNormalized[j][0], boundariesNormalized[k-1][1]+1);
					XSegment xsegment_match = new XSegment(reference_xsegments.get(j).getSource_file_name(), seg_range, str, strNormalized, new ArrayList());
					similar_reference_segments.add(xsegment_match);
					scores.add(prev_scores);
					end = (k-1);
					
					String trackedString = "";
					if(prev_scores[1] == Partial_match_score){
						//trackedString = createHtmlForPartialExactMatch(source_xsegment, xsegment, 0);
						trackedString = createHtmlForFuzzyMatch(str, text);
					}else{
						trackedString = createHtmlForFuzzyMatch(str, text);
					}
					diff_texts.add(trackedString);
				}
				
			}else{
				
				end = Math.max(end, j);
			}
			
		}
	}
	
	private void findReferenceMatches_FAST(XSegment xsegment, List<XSegment> reference_xsegments, String refStr, int[][] boundaries, String refStrNormalized, int[][] boundariesNormalized, int[] match_conclusions, int i, List<XSegment> similar_reference_segments, List<String> diff_texts, List<double[]> scores, HashSet<Integer> idx_exact_match, String language){
		
		String text = xsegment.getContent();
		String normalized_text = xsegment.getNormalized_content();

		double[] prev_scores = new double[2];
		
		int head = 0;
		for(int end = 0; end < reference_xsegments.size();){

			double[] cur_scores = getMatchScore(refStrNormalized, normalized_text, boundariesNormalized[head][0], boundariesNormalized[end][1], null, null, language);
			
			if(cur_scores[0] < prev_scores[0]){
				
				end--;
				double[] prev_scores2 = prev_scores;
				
				while(head < end){
					
					double[] cur_scores2 = getMatchScore(refStrNormalized, normalized_text, boundariesNormalized[head+1][0], boundariesNormalized[end][1], null, null, language);
					
					if(cur_scores2[0] < prev_scores2[0]){
						break;
					}else{
						prev_scores2 = cur_scores2;
						head++;
					}
				}
				
				if(prev_scores2[0] >= config.getThreshold_noraml()){
					
					match_conclusions[i] = Math.max(prev_scores2[0]<Partial_match_score?1:(prev_scores2[1]==Partial_match_score?2:3), match_conclusions[i]);
					
					if(!config.isShow_exact_match() && prev_scores2[0] == 100){
						continue;
					}
					
					if(!config.isShow_fuzzy_match() && prev_scores2[0] < 100 && prev_scores2[1] != Partial_match_score){
						continue;
					}
					
					if(!config.isShow_partial_exact_match() && prev_scores2[1] == Partial_match_score){
						continue;
					}
					
					if(prev_scores2[0] == 100){
						idx_exact_match.add(scores.size());
					}
					
					int s = Integer.parseInt(reference_xsegments.get(head).getSegment_id());
					String seg_range = (head==end)?(""+s):(s + " - " + (s+end-head));
					String str = refStr.substring(boundaries[head][0], boundaries[end][1]+1);
					String strNormalized = refStrNormalized.substring(boundariesNormalized[head][0], boundariesNormalized[end][1]+1);
					XSegment xsegment_match = new XSegment(reference_xsegments.get(head).getSource_file_name(), seg_range, str, strNormalized, new ArrayList());
					similar_reference_segments.add(xsegment_match);
					scores.add(prev_scores2);
					
					String trackedString = "";
					if(prev_scores2[1] == Partial_match_score){
						//trackedString = createHtmlForPartialExactMatch(source_xsegment, xsegment, 0);
						trackedString = createHtmlForFuzzyMatch(str, text);
					}else{
						trackedString = createHtmlForFuzzyMatch(str, text);
					}
					diff_texts.add(trackedString);
				}
				
				prev_scores = new double[2];
				head = end = end+1;
				
			}else{
				
				prev_scores = cur_scores;
				end++;
			}
			
		}
	}
	
	private void formatOrphanRow(Worksheet sheet, Cells cells, int row, boolean isUsed){
		
		Range range = cells.createRange(row, 0, 1, 2);
   		range.setColumnWidth(70.0D);
   		Style style = sheet.getWorkbook().createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		if(isUsed){
   			//style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 243, 203));
   		}
   		
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		StyleFlag styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//format indicator
   		range = cells.createRange(row, 2, 1, 2);
   		range.setColumnWidth(25.0D);
   		style = sheet.getWorkbook().createStyle();
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
	}
	
	private void formatOrphanHeaderRow(Worksheet sheet, Cells cells){
		
		Range range = cells.createRange(0, 0, 1, 4);
   		Style style = sheet.getWorkbook().createStyle();
   		style.getFont().setBold(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(237, 237, 237));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		StyleFlag styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
	}	
	
	private void formatSimilarityScoreCell(Cell cell, Cells cells) {
		
		if(cell.getValue() == null) return;
		String score = cell.getDisplayStringValue();
   		Range range = cells.createRange(cell.getRow(), cell.getColumn(), 1, 1);
		range.setColumnWidth(15.0D);
   		Style style = cell.getStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		if(score.equals("100")) {
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(0, 204, 153));
   		} else if(score.equals("100-")) {
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(0, 153, 204));
   		} else if(score.equals("0")) {
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 167, 167));
   		} else {
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(245, 254, 122));
   		}
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		StyleFlag styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
	}
	
	private void formatReferenceCell(Cell cell, Cells cells) {
		
   		Range range = cells.createRange(cell.getRow(), cell.getColumn(), 1, 1);
   		range.setColumnWidth(50.0D);
   		Style style = cell.getWorksheet().getWorkbook().createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 235, 156));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		StyleFlag styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
	}
	
	private void formatSourceIdCell(int start, int span, int src_size, int trg_size, Cells cells, boolean hasTarget) {
		
		//source segment id
		Range range = cells.createRange(start,0,span,1);
		range.setColumnWidth(15.0D);
   		Style style = cells.getStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		if(src_size != trg_size && this.isRunAgainstTargetReference && hasTarget){
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 199, 206));
   			Cell cell = cells.get(start, 0);
   			int idx = cell.getWorksheet().getComments().add(start, 0);
   			Comment comment = cell.getWorksheet().getComments().get(idx);
   			comment.setNote("source reference match: " + src_size + "\n" + "target reference match: " + trg_size);
   			comment.setHeight(40);
   			comment.setWidth(180);
   		}else{
   			style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		}
   		
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		StyleFlag styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
	}
	
	private LinkedHashMap<String, HashMap<String, List<String>>> consolidateReferenceFilesInfo(List<XSegment> refs, List<String> diffs, List<double[]> scores){
		
		LinkedHashMap<String, HashMap<String, List<String>>> map = new LinkedHashMap<String, HashMap<String, List<String>>>();
		for(int i = 0; i < refs.size(); i++){
			XSegment ref_x_segment = refs.get(i);
			String text = ref_x_segment.getContent();
			String normalized_text = ref_x_segment.getNormalized_content();
			String diff = diffs.get(i);
			String score = scores.get(i)[1] == Partial_match_score? "100-":Integer.toString((int)scores.get(i)[0]);
			String key = text + "#%#" + diff + "#%#" + score + "#%#" + normalized_text;
			if(!map.containsKey(key)){
				map.put(key, new HashMap<String, List<String>>());
			}
			HashMap<String, List<String>> metas = map.get(key);
			String filename = ref_x_segment.getSource_file_name();
			String id = ref_x_segment.getSegment_id();
			if(!metas.containsKey(filename)){
				metas.put(filename, new ArrayList<String>());
			}
			metas.get(filename).add(id);
		}
		
		return map;
	}
	
	private void formatExcelSheet(Worksheet sheet, int cell_start_index) {
		
		Workbook wb = sheet.getWorkbook();
		Cells cells = sheet.getCells();

		//source segment id title
		Range range = cells.createRange(0,0,1,1);
		range.setColumnWidth(15.0D);
		Style style = wb.createStyle();
		style.getFont().setBold(true);
		style.setTextWrapped(true);
		style.setHorizontalAlignment(TextAlignmentType.CENTER);
		style.setVerticalAlignment(TextAlignmentType.CENTER);
		style.setPattern(1);
		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
		StyleFlag styleFlag = new StyleFlag();
		styleFlag.setAll(true);
		range.applyStyle(style, styleFlag);

   		
   		//source segment text title
   		range = cells.createRange(0,1,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(198, 239, 206));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source segment text
   		range = cells.createRange(1,1,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(198, 239, 206));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source reference text title
   		range = cells.createRange(0,2,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 235, 156));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source reference text
   		range = cells.createRange(1,2,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 235, 156));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//track change title
   		range = cells.createRange(0,3,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//track changes
   		range = cells.createRange(1,3,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setWrapText(true);
   		styleFlag.setHorizontalAlignment(true);
   		styleFlag.setVerticalAlignment(true);
   		styleFlag.setIndent(true);
   		styleFlag.setBorders(true);
   		styleFlag.setLocked(true);
   		//styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//similarity score title
   		range = cells.createRange(0,4,1,1);
		range.setColumnWidth(15.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		
   		//source reference notes title
   		range = cells.createRange(0,5,1,1);
		range.setColumnWidth(30.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source reference notes
   		range = cells.createRange(1,5,cell_start_index-1,1);
		range.setColumnWidth(30.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source reference file name title
   		range = cells.createRange(0,6,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//source reference file name
   		range = cells.createRange(1,6,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		
   		//target
   		
   		//target segment text title
   		range = cells.createRange(0,7,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(198, 239, 206));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target segment text
   		range = cells.createRange(1,7,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(198, 239, 206));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target reference text title
   		range = cells.createRange(0,8,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 235, 156));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target reference text
   		range = cells.createRange(1,8,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(255, 235, 156));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//track change title
   		range = cells.createRange(0,9,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setWrapText(true);
   		styleFlag.setHorizontalAlignment(true);
   		styleFlag.setVerticalAlignment(true);
   		styleFlag.setIndent(true);
   		styleFlag.setBorders(true);
   		//styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//track changes
   		range = cells.createRange(1,9,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//similarity score title
   		range = cells.createRange(0,10,1,1);
		range.setColumnWidth(15.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		
   		//target reference notes title
   		range = cells.createRange(0,11,1,1);
		range.setColumnWidth(30.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target reference notes
   		range = cells.createRange(1,11,cell_start_index-1,1);
		range.setColumnWidth(30.0D);
   		style = wb.createStyle();
   		style.setLocked(false);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target reference file name title
   		range = cells.createRange(0,12,1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.getFont().setBold(true);
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.CENTER);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);
   		
   		//target reference file name
   		range = cells.createRange(1,12,cell_start_index-1,1);
		range.setColumnWidth(50.0D);
   		style = wb.createStyle();
   		style.setTextWrapped(true);
   		style.setHorizontalAlignment(TextAlignmentType.LEFT);
   		style.setIndentLevel(1);
   		style.setVerticalAlignment(TextAlignmentType.CENTER);
   		style.setPattern(1);
   		style.setForegroundColor(com.aspose.cells.Color.fromArgb(224, 224, 224));
   		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THIN,com.aspose.cells.Color.getBlack());
   		styleFlag = new StyleFlag();
   		styleFlag.setAll(true);
   		range.applyStyle(style, styleFlag);

   		sheet.freezePanes(1, 0, 1, 30);
	}
	
	private boolean isComparable(String s) {
		
		if(config.isIgnore_non_leveragable() && !this.leverageSupport.isLeveragable(s)) {
			return false;
		}
		
		if(config.isIgnore_number() && this.leverageSupport.isNumber(s)) {
			return false;
		}
		
		return true;
	}
	
	private double[] getMatchScore(String str1, String str2, Integer s1, Integer e1, Integer s2, Integer e2, String language){
		
		//this.count++;
		s1 = s1==null?0:s1;
		e1 = e1==null?(str1.length()-1):e1;
		s2 = s2==null?0:s2;
		e2 = e2==null?(str2.length()-1):e2;
		
		double score = (double)LevenshteinDistance.getMatchScore(str1,str2,s1,e1,s2,e2);
		
		String regex_1, regex_2;
		//if(s2 == -1 || e2 == -1) System.out.println("\"" + str1 + "\"" + "\"" + str2 + "\"");
		str1 = str1.substring(s1, e1+1);
		str2 = str2.substring(s2, e2+1);
		//if is character based language, use character level match, otherwise use word level match
		if(Locale.makeLocale(language).isFarEast()){
			regex_1 = Pattern.quote(str2);
			regex_2 = Pattern.quote(str1);
		}else{
			regex_1 = "(^|[ \\,\\.\\?:;\\!])" + Pattern.quote(str2) + "($|[ \\,\\.\\?:;\\!])";
			regex_2 = "(^|[ \\,\\.\\?:;\\!])" + Pattern.quote(str1) + "($|[ \\,\\.\\?:;\\!])";
		}
		
		Pattern p1 = Pattern.compile(regex_1);
		Matcher m1 = p1.matcher(str1);
		Pattern p2 = Pattern.compile(regex_2);
		Matcher m2 = p2.matcher(str2);
		
		if((str1.length() > 1 && str2.length() > 1) 
				&& (!str1.equals(str2)) 
					&& (m1.find() || m2.find())
						&& (Math.min(str1.length(), str2.length())*100/Math.max(str1.length(), str2.length())) >= config.getThreshold_partialmatch()) 
			return new double[]{score, Partial_match_score};
		
		return new double[]{score, 0};
	}
	
	private String createHtmlForFuzzyMatch(String olds, String news) {
		diff_match_patch dmp = new diff_match_patch();
		dmp.Diff_EditCost = 6;
		LinkedList<diff_match_patch.Diff> Diffs = dmp.diff_main(olds, news);
		dmp.diff_cleanupSemantic(Diffs);

		String trackedString = dmp.diff_prettyHtml(Diffs)
				.replaceAll("<([/\\s]*)ins([^>]*?)>", "<$1u>")
				.replaceAll("<([/\\s]*)del([^>]*?)>", "<$1s>")
				.replace("<u>", "<font color=\"#00b0f0\"><b><u>")
				.replace("</u>", "</u></b></font>")
				.replace("<s>", "<font color=\"#FF0000\"><b><s>")
				.replace("</s>", "</s></b></font>");
		
		return trackedString;
	}
	
	private String createHtmlForPartialExactMatch(XSegment src, XSegment ref, int flag) {
		if(src.getNormalized_content().length() > ref.getNormalized_content().length()) {
			return createHtmlForPartialExactMatch(ref, src, flag^1);
		}
		
		String res = ref.getNormalized_content();
		String bematchs = ref.getNormalized_content();
		String matchs = src.getNormalized_content();
		
		String color = flag == 1?"#00b0f0":"#ff0000";
		String format_s = flag == 1?"<u>":"<s>";
		String format_e = flag == 1?"</u>":"</s>";
		
		Matcher m = Pattern.compile("(?i)"+Pattern.quote(matchs)).matcher(bematchs);
		int lastend = 0;
		int offset = 0;
		while(m.find()) {
			if(lastend+offset != m.start()+offset){
				res = res.substring(0, lastend+offset) + "<font color=\"" + color + "\"><b>" + format_s + res.substring(lastend+offset, m.start()+offset) + format_e + "</b></font>" + res.substring(m.start()+offset, res.length());
				offset += 43;
			}
			lastend = m.end();
		}
		if(lastend+offset != res.length()) {
			res = res.substring(0, lastend+offset) + "<font color=\"" + color + "\"><b>" + format_s + res.substring(lastend+offset, res.length()) + format_e + "</b></font>";
		}
		//restore origianl string from normalized string and keep the html formatting
		//res = restoreUnnormalizedString(res, ref.getContent());
		
		return res;
	}
	
	/*private String restoreUnnormalizedString(String norm_s, String org_s){
		int i = 0;
		int j = 0;
		boolean ishtmltag = false;
		StringBuffer sb = new StringBuffer();
		while(i < norm_s.length()){
			char c = norm_s.charAt(i);
			if(c == '<'){
				ishtmltag = true;
			}
			
			if(!ishtmltag){
				if(c == ' '){
					while(org_s.charAt(j) == ' '){
						sb.append(org_s.charAt(j));
						j++;
					}
				}else{
					sb.append(org_s.charAt(j));
					j++;
				}
			}else{
				sb.append(c);
			}
			
			if(c == '>'){
				ishtmltag = false;
			}
			
			i++;
		}
		
		return sb.toString();
	}*/
	
	private <T> void removeAllOtherElementExcept(List<T> elements, HashSet<Integer> ids) {
		for(int i = elements.size()-1; i >= 0; i--){
			if(!ids.contains(i)){
				elements.remove(i);
			}	
		}
	}
	
	public void makeDummyTM(String tmpath, String sourcelanguage, String targetlanguage) throws IOException
	{    
		DateFormat dateFormat2 = new SimpleDateFormat("yyyyMMdd~HHmmss");
		Date date2 = new Date();
		String TNow2 = dateFormat2.format(date2);
		String dummyTMHeader = "%" + TNow2 + "\t" + "%User System\t%TU=00000000\t%" + sourcelanguage + "\t%" + "Wordfast TM v.546/00\t%" + targetlanguage + "\t%-----------\t\t\t\t\t                                                            .\r\n";

		Writer out = new OutputStreamWriter(new FileOutputStream(tmpath), "UTF-8");
		out.write(dummyTMHeader);
		out.close();

	}
  
	public void updateTM(String tmpath, List<String> txmls)
	{
		Configuration config = new BaseConfiguration();
    
		String tmLogin = "localtm://" + tmpath.replaceAll(",", "\\\\,");
		config.setProperty("cleaner.tm.loginurl", tmLogin);
		config.setProperty("cleaner.cleantosource", "false");
    	config.setProperty("cleaner.skipoutputfile", "true");
    	config.setProperty("cleaner.tm.throwexceptionOnEmptyTarget", "false");
    	config.setProperty("cleaner.writeUnconfirmedTUs", "true");
    
    	TxmlCleaner cleaner = new TxmlCleaner();
    	cleaner.setConfiguration(config);
    	cleaner.cleanupFiles(txmls);
	}
	
	public LinkedHashMap<String, String[]> AutoAlignSourceAndTargetReference(List<XSegment> source_reference_xsegments, List<XSegment> target_reference_xsegments, String sourcelanguage, String targetlanguage) throws Exception{
		
		LinkedHashMap<String, String[]> results = new LinkedHashMap<String, String[]>();
		
		String nbalignerfolder = tempfolder + File.separator + "nbaligner";
	    if (!new File(nbalignerfolder).exists()) {
	      new File(nbalignerfolder).mkdir();
	    }
	    
	    org.apache.commons.io.FileUtils.cleanDirectory(new File(nbalignerfolder));
	    
	    sourcelanguage = sourcelanguage.split("-")[0];
	    String nbsourcefolder = nbalignerfolder + File.separator + sourcelanguage;
	    new File(nbsourcefolder).mkdir();
	    org.dom4j.Document nbsource = DocumentHelper.createDocument();
	    org.dom4j.Element root_src = nbsource.addElement("txml");
	    root_src.addAttribute("locale", sourcelanguage);
	    root_src.addAttribute("version", "1.0");
	    root_src.addAttribute("segtype", "sentence");
	    org.dom4j.Element translatable_src = root_src.addElement("translatable");
	    translatable_src.addAttribute("blockId", "0");
	    
	    String nbtargetfolder = nbalignerfolder + File.separator + targetlanguage;
	    new File(nbtargetfolder).mkdir();
	    org.dom4j.Document nbtarget = DocumentHelper.createDocument();
	    org.dom4j.Element root_trg = nbtarget.addElement("txml");
	    root_trg.addAttribute("locale", targetlanguage);
	    root_trg.addAttribute("version", "1.0");
	    root_trg.addAttribute("segtype", "sentence");
	    org.dom4j.Element translatable_trg = root_trg.addElement("translatable");
	    translatable_trg.addAttribute("blockId", "0");
	    
	    int segmentId = 0;
	    for(XSegment segment : source_reference_xsegments){
	    	
	    	org.dom4j.Element segment_src = translatable_src.addElement("segment");
	        segment_src.addAttribute("segmentId", Integer.toString(segmentId));   
	        segmentId++;
	        segment_src.addElement("source").setContent(segment.getContent_list());
	    }
	    
	    segmentId = 0;
	    for(XSegment segment : target_reference_xsegments){
	    	
	    	org.dom4j.Element segment_trg = translatable_trg.addElement("segment");
	    	segment_trg.addAttribute("segmentId", Integer.toString(segmentId));   
	        segmentId++;
	        segment_trg.addElement("source").setContent(segment.getContent_list());
	    }
	    
	    OutputStreamWriter writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(nbsourcefolder + File.separator + sourcelanguage + ".txml")), "UTF8");
	    nbsource.write(writer);
	    writer.close();
	    
	    writer = new OutputStreamWriter(new BufferedOutputStream(new FileOutputStream(nbtargetfolder + File.separator + targetlanguage + ".txml")), "UTF8");
	    nbtarget.write(writer);
	    writer.close();
	    
	    
	    
	    String pahtexe = "\\\\10.2.50.190\\AutoAlignerCLI\\AutoAlignerCLI.exe";

	    ProcessBuilder pb = new ProcessBuilder(new String[] { pahtexe, "-i", nbalignerfolder, "-o", nbalignerfolder, "-lang_pairs", sourcelanguage + "_" + targetlanguage, "-lang_detect", "off", "-identicals", "-match_filenames", "-txml_or_xmx_output", "-docnames_output", "-tags", "-disallow_src_merging" });
	    pb.redirectErrorStream(true);
	    
	    Process p = pb.start();
	    InputStreamReader isr = new InputStreamReader(p.getInputStream());
	    BufferedReader br = new BufferedReader(isr);
	    
	    Writer auto_aligner_out = new OutputStreamWriter(new FileOutputStream(tempfolder + File.separator + "auto-aligner_output.txt"), "UTF-8");

	    String lineRead;
	    while ((lineRead = br.readLine()) != null)
	    {
		      sb.append(lineRead + "\n");
		      synchronized(auto_aligner_out){
		    	  auto_aligner_out.write(lineRead + "\r\n");
		      }
	    }
	    p.waitFor();
	    
	    auto_aligner_out.close();

	    for (File file : new File(nbalignerfolder).listFiles())
	    {
	    	if (file.getName().endsWith(".zip")) {
	    		UnzipFile.UnZipIt(file.getAbsolutePath(), nbalignerfolder);
	    	}
	    }
	    
	    String alignedtxml = "";
	    for (File file : new File(nbalignerfolder).listFiles())
	    {
	    	if (file.getName().endsWith(".txml")) {
	    		alignedtxml = file.getAbsolutePath();
	    	}
	    }
	    if (alignedtxml.equals("")) {
	    	throw new Exception("file wasn't aligned by nbaligner");
	    }
	    
	    
	    SAXReader reader = new SAXReader();
	    org.dom4j.Document alignedtxmldoc = reader.read(alignedtxml);
	    org.dom4j.Element root_alignedtxmldoc = alignedtxmldoc.getRootElement();
	    int no_source_index = 0;
	    for (int i = 0; i < root_alignedtxmldoc.elements("translatable").size(); i++)
	    {
	    	org.dom4j.Element translatable = (org.dom4j.Element)root_alignedtxmldoc.elements("translatable").get(i);
	    	for (int j = 0; j < translatable.elements("segment").size(); j++)
	    	{
	    		org.dom4j.Element segment = (org.dom4j.Element)translatable.elements("segment").get(j);
	    		org.dom4j.Element source = segment.element("source");
	    		org.dom4j.Element target = segment.element("target");
	    		
	    		String sourcetext = "";
	    		if(source != null && !source.getTextTrim().equals("")){
	    			sourcetext = FileUtils.reformatSegmentTextWithTags(source, true);
	    		}
	    		String targettext = "";
	    		if(target != null && !target.getTextTrim().equals("")){
	    			targettext = target.getTextTrim();
	    		}
	    		String score = "0";
	    		if(target != null && target.attribute("score") != null && !target.attribute("score").getValue().equals("0")){
	    			score = target.attribute("score").getValue();
	    		}
	    		
	    		if(sourcetext.equals("") && !targettext.equals("")){
	    			results.put(NULL_SOURCE_REFERENCE_STR + " - " + no_source_index, new String[] { targettext, "0" });
	    			no_source_index++;
	    		}else if(!sourcetext.equals("") && targettext.equals("")){
	    			if(!results.containsKey(sourcetext)){
	    				results.put(sourcetext, new String[] { "", "0" });
	    			}else{
	    				results.put(sourcetext+"[repeat]", new String[] { "", "0" });
	    			}
	    		}else if(!sourcetext.equals("") && !targettext.equals("")){
	    			if(score.equals("0")){
	    				if(!results.containsKey(sourcetext)){
	    					results.put(sourcetext, new String[] { "", "0" });
	    				}else{
	    					results.put(sourcetext+"[repeat]", new String[] { "", "0" });
	    				}
	    				results.put(NULL_SOURCE_REFERENCE_STR + " - " + no_source_index, new String[] { targettext, "0" });
		    			no_source_index++;
	    			}else{
	    				if(!results.containsKey(sourcetext)){
	    					results.put(sourcetext, new String[] { targettext, score });
	    				}else{
	    					results.put(sourcetext+"[repeat]", new String[] { targettext, score });
	    				}
	    			}
	    		}
	    	}
	    }
	    
	    return results;
	}
}
