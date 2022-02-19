package excelOperations;

import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.util.ArrayList;



public class merginExcel {
	
	public static boolean compare(File a, File b) throws IOException {
		ArrayList<String> f=execution.header(a);
		ArrayList<String> s=execution.header(b);
		return f.equals(s);
	}
	
	public static void main(String[] args) throws IOException 
	{
		File directory = new File("./data");
		File[] files = directory.listFiles();
		int k=0;
		//writing w = new writing();
		for(int i=0;i<files.length;i++)
		{
			for(int j=i+1;j<files.length;j++)
			{
				if(compare(files[i],files[j])) {
					
					writing.merge(files[i],files[j],k);
					k++;
				}
			}
		}
	}

}
