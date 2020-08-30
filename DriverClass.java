package AdvanceSelenium;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DriverClass {
	static FileOutputStream fout;
	

	public static void main(String[] args) throws IOException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		// TODO Auto-generated method stub

		GmailMethodClass gm=new GmailMethodClass();
		Method m[]=gm.getClass().getMethods();
		
		FileInputStream fis=new FileInputStream("‪‪C:\\Users\\RUHI\\Desktop\\Gmail.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet rsh1=wb.getSheetAt(0);
		int nur1=rsh1.getPhysicalNumberOfRows();
		
		XSSFSheet rsh2=wb.getSheetAt(0);
		int nur2=rsh2.getPhysicalNumberOfRows();
		int nuc2=rsh2.getRow(0).getPhysicalNumberOfCells();
		
		for(int i=1;i<nur1;i++) {
			String runmode=rsh1.getRow(i).getCell(2).getStringCellValue();
			String tid = rsh1.getRow(i).getCell(0).getStringCellValue();
			
			if(runmode.equalsIgnoreCase("yes")){
				for(int j=1;j<nur2;j++) {
					String sid=rsh2.getRow(j).getCell(0).getStringCellValue();
					
					if(tid.equalsIgnoreCase(sid)){
						String method=rsh2.getRow(j).getCell(2).getStringCellValue();
						String l=rsh2.getRow(j).getCell(3).getStringCellValue();
						String d=rsh2.getRow(j).getCell(4).getStringCellValue();
						String c=rsh2.getRow(j).getCell(5).getStringCellValue();
						
						    for(int k=0;k<m.length;k++) {
						    	if( method.equalsIgnoreCase(m[k].getName())){
						    		String res=(String)m[k].invoke(gm, l,d,c);
						    		XSSFCell cell=rsh2.getRow(j).getCell(6);
						    		cell.setCellValue(res);
	fout = new FileOutputStream("‪C:\\Users\\RUHI\\Desktop\\gmailexcel1.xlsx");
						    		
						    	}
						    }
						
					}
					
				}
			}
		}
		
		wb.write(fout);	
		wb.close();	
		
		
	}

}
