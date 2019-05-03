package com.desktoptool.xcompare;

import java.util.List;

public class XCompareProgress implements Runnable{
	
	private List<String> progresses;
	
	public XCompareProgress(List<String> progresses){
		this.progresses = progresses;
	}
	
	public void run(){
		
		StringBuffer sb = new StringBuffer();
		for(int i = 0; i < this.progresses.size(); i++){
			sb.append(progresses.get(i) + (i!=this.progresses.size()-1?" | ":""));
		}
		System.out.print(sb.toString() + "\r");
	}
}
