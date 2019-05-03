package com.desktoptool.xcompare;

public class LevenshteinDistance 
{
	private static int minimum(int a, int b, int c)
    {
        return Math.min(Math.min(a, b), c);
    }

    public static int computeLevenshteinDistance(CharSequence str1, CharSequence str2, int s1, int e1, int s2, int e2)
    {
    	int len1 = e1 - s1 + 1;
    	int len2 = e2 - s2 + 1;
        int[][] distance = new int[len1 + 1][len2 + 1];

        for (int i = 0; i <= len1; i++)
            distance[i][0] = i;
        for (int j = 1; j <= len2; j++)
            distance[0][j] = j;

        for (int i = 1; i <= len1; i++)
            for (int j = 1; j <= len2; j++)
                distance[i][j] = minimum(distance[i - 1][j] + 1, 
                                         distance[i][j - 1] + 1, 
                                         distance[i - 1][j - 1] + ((str1.charAt(i + s1 - 1) == str2.charAt(j + s2 - 1)) ? 0 : 1));

        return distance[len1][len2];
    }
    
    public static double getMatchScore(CharSequence str1, CharSequence str2, int s1, int e1, int s2, int e2)
    {
    	double distance = computeLevenshteinDistance(str1, str2, s1, e1, s2, e2);
    	int maxLen = Math.max(e1-s1+1, e2-s2+1);
    	return ((((double)maxLen - distance)/maxLen) * 100);
    }
}
