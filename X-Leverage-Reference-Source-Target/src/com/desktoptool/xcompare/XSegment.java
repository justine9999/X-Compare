package com.desktoptool.xcompare;

import java.util.List;

import org.dom4j.Element;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class XSegment {
	
	private String source_file_name;
	private String segment_id;
	private String content;
	private String normalized_content;
	private List content_list;
	
	public XSegment(String source_file_name, String segment_id, String content, String normalized_content, List content_list) {
		super();
		this.source_file_name = source_file_name;
		this.segment_id  = segment_id;
		this.content = content;
		this.normalized_content = normalized_content;
		this.content_list = content_list;
	}
}
