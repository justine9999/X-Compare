package com.desktoptool.xcompare;

import java.util.List;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class XRecord {
	
	private String segment_id;
	private String content;
	List<XRecordItem> items;
	
	public XRecord(String segment_id, String content, List<XRecordItem> items) {
		super();
		this.segment_id = segment_id;
		this.content = content;
		this.items = items;
	}
	
	
}
