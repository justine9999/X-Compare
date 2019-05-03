package com.desktoptool.xcompare;

public enum MatchStatus {

	NO_MATCH(0),
    FUZZY_MATCH(1),
    PARTIAL_EXACT_MATCH(2),
	EXACT_MATCH(3);
	
	private final int value;
	
	private MatchStatus(int value) {
        this.value = value;
    }
}
