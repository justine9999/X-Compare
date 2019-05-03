package com.desktoptool.xcompare;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.Serializable;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

public class UnzipFile
  implements Serializable
{
	private static final long serialVersionUID = 1L;

public static void UnZipIt(String zipFile, String outputFolder)
    throws Exception
  {    
    byte[] buffer = new byte[1024];
    
    File folder = new File(outputFolder);
    if (!folder.exists()) {
      folder.mkdir();
    }
    ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile));
    ZipEntry ze = zis.getNextEntry();
    while (ze != null)
    {
      String fileName = ze.getName();
      File newFile = new File(outputFolder + File.separator + fileName);
      if (ze.isDirectory())
      {
        new File(newFile.getParent()).mkdirs();
      }
      else
      {
        new File(newFile.getParent()).mkdirs();
        FileOutputStream fos = new FileOutputStream(newFile);
        int len;
        while ((len = zis.read(buffer)) > 0) {
          fos.write(buffer, 0, len);
        }
        fos.close();
      }
      ze = zis.getNextEntry();
    }
    zis.closeEntry();
    zis.close();
  }
}
