package com.desktoptool.xcompare;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.gs4tr.foundation.core.commandline.Arguments;
import org.gs4tr.foundation.locale.Locale;

public class XCompareRSTBatch {

	public static void main(String[] argv) {
		
		System.exit(main2(argv));
	}

	public static int main2(String[] args) {
		
		final int MAX_THREAD_COUNT = 8;
		final int TIME_OUT_HOUR = 3;
		
		Arguments getopt = new Arguments();
		String parameter = "";
		String configurationfile = "";
	      
		getopt.setUsage(new String[] { "Usage: java XCompareRSTBatch [-c config file] [-p parameter]", "", "X-Compare Reference Soure Target Batch Mode" });
		getopt.parseArgumentTokens(args, new char[] { 'c', 'p' });
		int c;
		while ((c = getopt.getArguments()) != -1) {
			switch (c)
			{
				case 112: 
					parameter = getopt.getStringParameter();
					if (parameter.equals("")) {
						System.err.println("invalid argument for -p");
						System.out.println();
						pressAnyKeyToContinue();
						return 1;
					}
					break;
				case 99: 
					configurationfile = getopt.getStringParameter();
					if (!new File(configurationfile).exists()) {
						System.err.println("invalid argument for -c");
						System.out.println();
						pressAnyKeyToContinue();
						return 1;
					}
					break;
			}
		}
		
		/*String folder = "C:\\Users\\mxiang\\work\\4023\\2";
		String config = "N:\\j\\folders\\jx'\\XCompareRST\\XCompare_config.ini";
		
		String srcstr = "";
		String refstr = "";
		String finalstr = "";
		
		srcstr = new File(folder + File.separator + "src").listFiles()[0].getAbsolutePath();
		refstr = new File(folder + File.separator + "ref").listFiles()[0].getAbsolutePath();
		for(File file : new File(folder + File.separator + "tref").listFiles()){
			String trefstr = file.getAbsolutePath();
			String srclang = "EN";
			String trglang = FileUtils.getNameWithoutExtension(file.getName());
			String jobstr = srcstr + "*" + refstr + "*" + trefstr + "*" + srclang + "*" + trglang + ";";
			finalstr += jobstr;
		}
		finalstr = finalstr.substring(0, finalstr.length()-1);
		finalstr = "java -jar XCompare-Batch.jar -c \"" + config + "\" -p \"" + finalstr + "\"";
		System.out.println(finalstr);
		if(true){
			return 1;
		}*/

		
		//parameter = "N:\\j\\folders\\jx'\\XCompareRST\\file\\src\\src.txml*N:\\j\\folders\\jx'\\XCompareRST\\file\\ref\\ref.docx*N:\\j\\folders\\jx'\\XCompareRST\\file\\tref\\tref.xlsx*EN*ES";
		//parameter = "N:\\j\\folders\\jx'\\XCompareRST\\batch\\job1\\src\\ACT_ESES.xlsx.txml*N:\\j\\folders\\jx'\\XCompareRST\\batch\\job1\\ref\\REF_EN_ACT_DE.doc.txml|N:\\j\\folders\\jx'\\XCompareRST\\batch\\job1\\ref\\REF_EN_ACT_DE.doc.txml.doc.txml**EN-US*ES-ES";
		//configurationfile = "N:\\j\\folders\\jx'\\XCompareRST\\XCompare_config.ini";
		
		ApplyAsposeLicenseToWord();
		BasicConfigurator.configure();
		Logger.getRootLogger().setLevel(Level.OFF);
		
		String tempfolder = System.getProperty("java.io.tmpdir") + File.separator + "x-compare";
		if(!new File(tempfolder).exists()){
			new File(tempfolder).mkdir();
		}
		
		try {
			org.apache.commons.io.FileUtils.cleanDirectory(new File(tempfolder));
		} catch (IOException e) {
			System.err.println("cannot clean temporary folder, maybe it's open or being used by other process!");	
			e.printStackTrace();
			System.out.println();
			pressAnyKeyToContinue();
			return 1;
		}
		
		ExecutorService executorService = Executors.newFixedThreadPool(MAX_THREAD_COUNT);
		ScheduledExecutorService executorService_progress = Executors.newSingleThreadScheduledExecutor();
		
		String[] jobs = parameter.split(";");
		if(jobs.length == 0){
			System.err.println("no job found in the parameter");
			System.out.println();
			pressAnyKeyToContinue();
			return 1;
		}
		
		Writer out = null;
		try {
			out = new OutputStreamWriter(new FileOutputStream(tempfolder + File.separator + "log.txt"), "UTF-8");
		} catch (Exception ex) {
			try {
				out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
			ex.printStackTrace();
			pressAnyKeyToContinue();
			return 1;
		}
		
		List<XCompareJob> tasks = new ArrayList<XCompareJob>();
		String[] p_array = new String[jobs.length];
		Arrays.fill(p_array, "S");
		List<String> progresses = Arrays.asList(p_array);
		
		for(int i = 0; i < jobs.length; i++){
			
			String job = jobs[i];
			String[] fields = job.split("\\*");
			if(fields.length != 5){
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
				System.err.println("invalid parameter for job " + i);
				System.out.println();
				pressAnyKeyToContinue();
				return 1;
			}
			String[] srcs = fields[0].split("\\|");
			List<File> sourcefiles = new ArrayList<File>();
			for(String src : srcs){
				sourcefiles.add(new File(src));
			}
			String[] srefs = fields[1].split("\\|");
			List<File> sourcereferencefiles = new ArrayList<File>();
			for(String sref : srefs){
				sourcereferencefiles.add(new File(sref));
			}
			String[] trefs = fields[2].split("\\|");
			List<File> targetreferencefiles = new ArrayList<File>();
			for(String tref : trefs){
				targetreferencefiles.add(new File(tref));
			}
			String sourcelanguage = fields[3].toUpperCase();
			try{
				Locale.makeLocale(sourcelanguage);
			}catch(Exception ex){
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
				System.err.println("invalid source language code for job " + i + ": " + sourcelanguage);
				System.out.println();
				pressAnyKeyToContinue();
				return 1;
			}
			String targetlanguage = fields[4].toUpperCase();
			try{
				Locale.makeLocale(targetlanguage);
			}catch(Exception ex){
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
				System.err.println("invalid target language code for job " + i + ": " + targetlanguage);
				System.out.println();
				pressAnyKeyToContinue();
				return 1;
			}
			
			String jobtempfolder = tempfolder + File.separator + "job_" + i;
			if(!new File(jobtempfolder).exists()){
				new File(jobtempfolder).mkdir();
			}
			
			XCompareJob xCompareJob = new XCompareJob(i, sourcelanguage, targetlanguage, sourcefiles, sourcereferencefiles, targetreferencefiles, jobtempfolder, configurationfile, progresses);
			tasks.add(xCompareJob);
		}
		
		System.out.println("[" + jobs.length + " jobs collected, starting jobs]");
		System.out.println();
		
		System.out.println("Start[S]  Convert[C]  Align[A]  Fail[F]");
		System.out.println();
		
		XCompareProgress xCompareProgress = new XCompareProgress(progresses);
		executorService_progress.scheduleAtFixedRate(xCompareProgress, 0, 10, TimeUnit.MILLISECONDS);
		
		try {
			List<Future<String>> results = executorService.invokeAll(tasks);
			executorService.shutdown();
			executorService.awaitTermination(TIME_OUT_HOUR, TimeUnit.HOURS);

			for(int i = 0; i < results.size(); i++){
				Future<String> result = results.get(i);
				try {
					out.write("##################### Job " + i + " #####################" + "\n\n");
					out.write(result.get() + "\n\n");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
			
		} catch (InterruptedException e) {
			System.err.println("parallel execution failed");
			e.printStackTrace();
		}
		
		try {
			out.close();
		} catch (IOException ex) {
			ex.printStackTrace();
		}
		
		executorService_progress.shutdownNow();
		
		System.out.println();
		System.out.println();
		System.out.println("[job finished]");
		pressAnyKeyToContinue();
		
		return 0;
	}
	
	public static void ApplyAsposeLicenseToWord()
	{
		com.aspose.cells.License license= new com.aspose.cells.License();
		try {
			System.out.println();
			System.out.println("[applying aspose license]");
			InputStream licence = XCompareRST.class.getClassLoader().getResourceAsStream("properties/aspose.license");
			license.setLicense(licence);
		}
		catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private static void pressAnyKeyToContinue()
	{ 
        System.out.println("[Press Enter key to close]");
        try
        {
            System.in.read();
        }  
        catch(Exception e)
        {}  
	}
}
