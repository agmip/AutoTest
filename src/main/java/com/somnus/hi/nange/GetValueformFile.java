// 8.23 NinthTest
package com.somnus.hi.nange;

import java.io.File;

public class GetValueformFile {
	public static String dirstr;
	
	public static int GetFileNum(String dir)
	{
		File d = new File(dir);
		File list[] = d.listFiles();
		return list.length;
	}
	
	public static File[] GetAllFiles(String dir)
	{
		System.out.println("The test folder is: " + dir);
		File d = new File(dir);
		File list[] = d.listFiles();

		for(int i=0; i<list.length; i++)
		{
			list[i].getAbsolutePath();
		}	
		return list;
	}	  
}
