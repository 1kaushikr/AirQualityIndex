package EXCEL;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import org.apache.poi.xssf.usermodel.*;
public class AQI {
	public static void main(String[] args) throws IOException {
		String path = ".\\DATA\\K.xlsx";
		FileInputStream Input = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(Input);
		XSSFSheet sheet = workbook.getSheet("Sheet1"); 
		XSSFWorkbook workbook1 = new XSSFWorkbook();
		XSSFSheet sheet1 = workbook1.createSheet("X");
		int rows = 365;
		int cols = 8;
		for (int r = 0; r < rows; r++)
		{
			XSSFRow row1 = sheet1.createRow(r);
			XSSFRow row = sheet.getRow(r+17);
			for(int c = 0; c < cols; c++)
			{ 
				if (c<4)
				{
					XSSFCell cell = row.getCell(c+2);
					XSSFCell cell1 = row1.createCell(c);
					double d = cell.getNumericCellValue();
					cell1.setCellValue(d);
				}
				if (c==4)
				{ 
					XSSFCell cell = row.getCell(c-2);
					XSSFCell cell1 = row1.createCell(c);
					double d = cell.getNumericCellValue();
					double e = PM2(d);
					cell1.setCellValue(e);
				}
				if (c==5)
				{ 
					XSSFCell cell = row.getCell(c-2);
					XSSFCell cell1 = row1.createCell(c);
					double d = cell.getNumericCellValue();
					double e = PM1(d);
					cell1.setCellValue(e);
				}
				if (c==6)
				{ 
					XSSFCell cell = row.getCell(c-2);
					XSSFCell cell1 = row1.createCell(c);
					double d = cell.getNumericCellValue();
					double e = N(d);
					cell1.setCellValue(e);
				}
				if (c==7)
				{ 
					XSSFCell cell = row.getCell(c-2);
					XSSFCell cell1 = row1.createCell(c);
					double d = cell.getNumericCellValue();
					double e = S(d);
					cell1.setCellValue(e);
				}
			}
		}
		for (int r = 0; r < rows; r++)
		{
			XSSFRow row1 = sheet1.getRow(r);
			XSSFCell cell1 = row1.getCell(4);
			XSSFCell cell2 = row1.getCell(5);
			XSSFCell cell3 = row1.getCell(6);
			XSSFCell cell4 = row1.getCell(7);
			XSSFCell cell5 = row1.createCell(8);
			double d = cell1.getNumericCellValue();double e = cell2.getNumericCellValue();
			double f = cell3.getNumericCellValue();
			double g = cell4.getNumericCellValue();
			double h = max1(d,e,f,g);
			cell5.setCellValue(h);
		}
		String Path1 = ".\\DATA\\opop.xlsx";
		FileOutputStream out = new FileOutputStream(Path1);
		workbook1.write(out);
		out.close(); 
	}
	public static double max1(double d1, double d2, double d3, double d4) {
		double k = d1;
		if(d2>d1) {
			k = d2;
		}
		if(d3>d1) {
			k = d3;
		}
		if(d4>d1) {
			k = d4;
		}
		return k;
	}
	public static double PM2(double d) {
		DecimalFormat df = new DecimalFormat("###.##");
		double f;
		if (0<=d && d<=30.5) {
			f = Double.parseDouble(df.format(((50.0-0.0)/(30.0-0.0))*(d-0.0)+0.0));
			return f;
		}
		else if (30.5<d && d<=60.5) {
			f = Double.parseDouble(df.format(((100.0-51.0)/(60.0-31.0))*(d-31.0)+51.0));
			return f;
		}
		else if (60.5<d && d<=90.5) {
			f = Double.parseDouble(df.format(((200.0-101.0)/(90.0-61.0))*(d-61.0) + 
					(101.0)));
			return f;
		}
		else if (90.5<d && d<=120.5) {
			f = Double.parseDouble(df.format(((300.0-201.0)/(120.0-91.0))*(d-91.0) + 
					(201.0)));
			return f;
		}
		else if (120.5<d && d<=250) {
			f = Double.parseDouble(df.format(((400.0-301.0)/(250.0-121.0))*(d-121.0) + 
					(301.0)));
			return f;
		}
		else
		{
			f = Double.parseDouble(df.format(500.000));
			return f;
		}
	}
	public static double PM1(double d) {
		DecimalFormat df = new DecimalFormat("###.##");
		double f;
		if (0<=d && d<=50.5) {
			f = Double.parseDouble(df.format(((50.0-0.0)/(50.0-0.0))*(d-0.0)+0.0));
			return f;
		}
		else if (50.5<d && d<=100.5) {
			f = Double.parseDouble(df.format(((100.0-51.0)/(100.0-51.0))*(d-51.0)+51.0));
			return f;
		}
		else if (100.5<d && d<=250.5) {
			f = Double.parseDouble(df.format(((200.0-101.0)/(250.0-101.0))*(d-101.0) + 
					(101.0)));
			return f;
		}
		else if (250.5<d && d<=350.5) {f = Double.parseDouble(df.format(((300.0-201.0)/(350.0-251.0))*(d-251.0) + 
				(201.0)));
		return f;
		}
		else if (350.5<d && d<=430.0) {
			f = Double.parseDouble(df.format(((400.0-301.0)/(430.0-351.0))*(d-351.0) + 
					(301.0)));
			return f;
		}
		else
		{
			f = Double.parseDouble(df.format(500.000));
			return f;
		}
	}
	public static double N(double d) {
		DecimalFormat df = new DecimalFormat("###.##");
		double f;
		if (0<=d && d<=40.5) {
			f = Double.parseDouble(df.format(((50.0-0.0)/(40.0-0.0))*(d-0.0)+0.0));
			return f;
		}
		else if (40.5<d && d<=80.5) {
			f = Double.parseDouble(df.format(((100.0-51.0)/(80.0-41.0))*(d-41.0)+51.0));
			return f;
		}
		else if (80.5<d && d<=180.5) {
			f = Double.parseDouble(df.format(((200.0-101.0)/(180.0-81.0))*(d-81.0) + 
					(101.0)));
			return f;
		}
		else if (180.5<d && d<=280.5) {
			f = Double.parseDouble(df.format(((300.0-201.0)/(280.0-181.0))*(d-181.0) + 
					(201.0)));
			return f;
		}
		else if (280.5<d && d<=400) {
			f = Double.parseDouble(df.format(((400.0-301.0)/(400.0-281.0))*(d-281.0) + 
					(301.0)));
			return f;
		}
		else
		{
			f = Double.parseDouble(df.format(500.000));
			return f;
		}
	}
	public static double S(double d) {
		DecimalFormat df = new DecimalFormat("###.##");
		double f;
		if (0<=d && d<=40.5) {
			f = Double.parseDouble(df.format(((50.0-0.0)/(40.0-0.0))*(d-0.0)+0.0));
			return f;
		}
		else if (40.5<d && d<=80.5) {
			f = Double.parseDouble(df.format(((100.0-51.0)/(80.0-41.0))*(d-41.0)+51.0));
			return f;
		}
		else if (80.5<d && d<=380.5) {
			f = Double.parseDouble(df.format(((200.0-101.0)/(380.0-81.0))*(d-81.0) + 
					(101.0)));
			return f;
		}
		else if (380.5<d && d<=800.5) {
			f = Double.parseDouble(df.format(((300.0-201.0)/(800.0-381.0))*(d-381.0) + 
					(201.0)));
			return f;
		}
		else if (800.5<d && d<=1600) {
			f = Double.parseDouble(df.format(((400.0-301.0)/(1600.0-801.0))*(d-801.0) + 
					(301.0)));
			return f;
		}
		else {
			f = Double.parseDouble
					(df.format(500.000));
			return f;
		}
	}

}