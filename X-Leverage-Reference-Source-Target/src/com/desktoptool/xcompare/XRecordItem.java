package com.desktoptool.xcompare;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class XRecordItem {

	private String content;
	private String tracked_content;
	private String score;
	private String notes;
	private String location;
	
	public XRecordItem(String content, String tracked_content, String score, String notes, String location) {
		super();
		this.content = content;
		this.tracked_content = tracked_content;
		this.score = score;
		this.notes = notes;
		this.location = location;
	}
}
