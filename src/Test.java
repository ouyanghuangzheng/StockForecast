import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;

public class Test {

	/**
	 * @param args
	 */

	static int RowNumber = 0;
	static int Difference2 = 0;
	static int Difference3 = 0;
	static List<double[]> list = new ArrayList<double[]>();

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			run();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	static void run() throws IOException {

		final double limit = 1.5;// you can change the limitation here

		// double[][] points = pointInsertUsingTxt();
		double[][] points = pointInsertUsingExcel();
		int j = linearRegression(limit, points);
		if (j != -1) {
			double[][] points0 = new double[2][j + 1];
			double[][] points1 = new double[2][points[0].length - j];

			for (int i = 0; i <= j; i++) {
				points0[0][i] = points[0][i];
				points0[1][i] = points[1][i];
			}
			for (int i = j; i <= points[0].length - 1; i++) {
				points1[0][i - j] = points[0][i];
				points1[1][i - j] = points[1][i];
			}

			connection(points[0][0], points[1][0], points[0][j], points[1][j],
					limit, points0);
			connection(points[0][j], points[1][j],
					points[0][points[0].length - 1],
					points[1][points[0].length - 1], limit, points1);
			output();
			// outputIndices();
			printIndices2();
		}
	}

	static double[][] pointInsertUsingTxt() throws IOException {

		BufferedReader input = new BufferedReader(new FileReader(
				"file\\Points.txt"));
		String s, text = new String();
		while ((s = input.readLine()) != null)
			text += s;
		input.close();
		String[] location = text.split("\\D");
		double[][] data = new double[2][location.length / 2];
		for (int i = 0; i < location.length / 2; i++) {
			data[0][i] = Double.parseDouble(location[2 * i]);
			data[1][i] = Double.parseDouble(location[2 * i + 1]);
		}
		for (int i = 0; i < data[0].length - 1; i++) {// sorting
			double min = data[0][i];
			int n = i;
			for (int j = i; j <= data[0].length - 1; j++) {
				if (data[0][j] < min) {
					n = j;
				}
			}
			if (n != i) {
				double d0 = data[0][i];
				double d1 = data[1][i];
				data[0][i] = data[0][n];
				data[1][i] = data[1][n];
				data[0][n] = d0;
				data[1][n] = d1;
			}
			// System.out.println(data[0][i]+"  "+data[1][i]);
		}

		return data;
	}

	static double[][] pointInsertUsingExcel() throws IOException {

		String fileToBeRead = "file\\data.xls";
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(
				fileToBeRead)); // 创建对Excel工作簿文件的引用
		HSSFSheet sheet = workbook.getSheet("Sheet2"); // 创建对工作表的引用
		int rows = sheet.getPhysicalNumberOfRows();// 获取表格的
		String text = "";
		for (int r = 0; r < rows; r++) { // 循环遍历表格的行
			HSSFRow row = sheet.getRow(r); // 获取单元格中指定的行对象
			if (row != null) {
				int cells = row.getPhysicalNumberOfCells();// 获取单元格中指定列对象
				for (short c = 0; c < cells; c++) { // 循环遍历单元格中的列
					HSSFCell cell = row.getCell((short) c); // 获取指定单元格中的列
					if (cell != null
							&& cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
						text += cell.getNumericCellValue() + ",";
					}
				}
			}
		}

		String[] location = text.split(",");
		double[][] data = new double[2][location.length / 2];
		for (int i = 0; i < location.length / 2; i++) {
			data[0][i] = Double.parseDouble(location[2 * i]);
			data[1][i] = Double.parseDouble(location[2 * i + 1]);
			// System.out.println(data[0][i]+"  "+data[1][i]);
		}
		for (int i = 0; i < data[0].length - 1; i++) {// sorting
			double min = data[0][i];
			int n = i;
			for (int j = i; j <= data[0].length - 1; j++) {
				if (data[0][j] < min) {
					n = j;
				}
			}
			if (n != i) {
				double d0 = data[0][i];
				double d1 = data[1][i];
				data[0][i] = data[0][n];
				data[1][i] = data[1][n];
				data[0][n] = d0;
				data[1][n] = d1;
			}
		}

		return data;
	}

	static int linearRegression(double limit, double[][] i) {

		int length = i[0].length;
		double adverage_x = 0;
		double adverage_y = 0;
		double k = 0;
		double b = 0;

		for (int j = 0; j < length; j++) {
			adverage_x += i[0][j];
		}
		adverage_x /= length;// final

		for (int j = 0; j < length; j++) {
			adverage_y += i[1][j];
		}
		adverage_y /= length;// final

		double k1 = 0;
		for (int j = 0; j < length; j++) {
			k1 += (i[0][j] - adverage_x) * (i[1][j] - adverage_y);
		}

		double k2 = 0;
		for (int j = 0; j < length; j++) {
			k2 += Math.pow((i[0][j] - adverage_x), 2);
		}

		k = k1 / k2;// final

		b = adverage_y - k * adverage_x;// final

		System.out.println("The linear regression is " + "y=" + k + "x+" + b);

		return max(k, b, limit, i);
	}

	static int max(double k, double b, double limit, double[][] i) {
		double max = 0;
		int number = -1;
		if (i[0].length <= 2) {
			return -1;
		}
		for (int j = 0; j < i[0].length; j++) {
			if (Math.abs(k * i[0][j] + b - i[1][j]) > max) {
				max = Math.abs(k * i[0][j] + b - i[1][j]);
				number = j;
			}
		}
		if (max * Math.pow(1 / (1 + Math.pow(k, 2)), 1 / 2) > limit) {
			return number;
		}
		return -1;
	}

	static void connection(double x1, double y1, double x2, double y2,
			double limit, double[][] i) {
		double k = (y1 - y2) / (x1 - x2);
		double b = y1 - k * x1;
		int j = max(k, b, limit, i);
		if (j != -1) {
			double[][] i0 = new double[2][j + 1];
			double[][] i1 = new double[2][i[0].length - j];

			for (int a = 0; a <= j; a++) {
				i0[0][a] = i[0][a];
				i0[1][a] = i[1][a];
			}
			for (int a = j; a <= i[0].length - 1; a++) {
				i1[0][a - j] = i[0][a];
				i1[1][a - j] = i[1][a];
			}

			connection(x1, y1, i[0][j], i[1][j], limit, i0);
			connection(i[0][j], i[1][j], x2, y2, limit, i1);
		} else {
			System.out.println("y=" + k + "x+" + b + " with points (" + x1
					+ "," + y1 + ") and (" + x2 + "," + y2
					+ ") is fit for limitation!");
			double[] group = new double[6];
			group[0] = k;
			group[1] = b;
			group[2] = x1;
			group[3] = y1;
			group[4] = x2;
			group[5] = y2;
			list.add(group);
		}
	}

	static void output() {

		HSSFWorkbook wb = new HSSFWorkbook();// 建立新HSSFWorkbook对象
		HSSFSheet sheet = wb.createSheet("dataoutput");// 建立新的sheet对象

		/**************************************************************************************/
		HSSFWorkbook wb2 = new HSSFWorkbook();
		HSSFSheet sheet2 = wb2.createSheet("dataoutput2");
		/**************************************************************************************/
		HSSFWorkbook wb3 = new HSSFWorkbook();
		HSSFSheet sheet3 = wb3.createSheet("dataoutput3");
		/**************************************************************************************/

		while (!list.isEmpty()) {
			HSSFRow row = sheet.createRow((short) (RowNumber++));// 建立新行 Create
																	// a row and
																	// put some
																	// cells in
																	// it. Rows
																	// are 0
																	// based.
			// HSSFCell cell0 = row.createCell((short)0);//建立新cell
			double[] group = list.remove(0);
			row.createCell((short) 0).setCellValue(group[0]);
			row.createCell((short) 1).setCellValue(group[1]);
			row.createCell((short) 2).setCellValue(group[2]);
			row.createCell((short) 3).setCellValue(group[3]);
			row.createCell((short) 4).setCellValue(group[4]);
			row.createCell((short) 5).setCellValue(group[5]);

			/**************************************************************************************/
			HSSFRow row2 = sheet2
					.createRow((short) (RowNumber - 1 + Difference2++));
			row2.createCell((short) 0).setCellValue(group[2]);
			row2.createCell((short) 1).setCellValue(0.5);
			HSSFRow row3 = sheet2
					.createRow((short) (RowNumber - 1 + Difference2));
			row3.createCell((short) 0).setCellValue((group[2] + group[4]) / 2);
			if (group[0] >= 0) {
				row3.createCell((short) 1).setCellValue(0);
			} else {
				row3.createCell((short) 1).setCellValue(1);
			}
			/**************************************************************************************/
			for (int i = (int) group[2]; i < (int) group[4]; i++) {
				HSSFRow row4 = sheet3
						.createRow((short) (RowNumber - 1 + Difference3++));
				row4.createCell((short) 0).setCellValue(i);
				double signal = 0;
				if (i - (int) group[2] <= ((int) group[4] - (int) group[2]) / 2) {
					signal = (i - group[2]) / ((group[4] - group[2]) / 2) * 0.5;
				} else {
					signal = (group[4] - i) / ((group[4] - group[2]) / 2) * 0.5;
				}
				if (group[0] >= 0) {
					signal = -signal;
				}
				signal = 0.5 + signal;
				row4.createCell((short) 1).setCellValue(signal);
			}
			Difference3--;
			/**************************************************************************************/
		}

		try {
			FileOutputStream fileOut = new FileOutputStream(
					"file\\dataOutput.xls");
			wb.write(fileOut);// 把Workbook对象输出到文件workbook.xls中
			fileOut.close();

			/**************************************************************************************/
			FileOutputStream fileOut2 = new FileOutputStream(
					"file\\dataOutput2.xls");
			wb2.write(fileOut2);
			fileOut2.close();
			/**************************************************************************************/
			FileOutputStream fileOut3 = new FileOutputStream(
					"file\\dataOutput3.xls");
			wb3.write(fileOut3);
			fileOut3.close();
			/**************************************************************************************/

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	static double[][] MA(int i) throws IOException {// 5102060

		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for (int j = i - 1; j < length; j++) {
			result[0][j] = data[0][j];
			for (int k = 0; k < i; k++) {
				result[1][j] += data[1][j - k];
			}
			result[1][j] /= i;
		}
		result[1][0] = i;

		return result;
	}

	static double[][] BIAS(int i) throws IOException {// 61224

		double[][] data = pointInsertUsingExcel();
		double[][] ma = MA(i);
		int length = data[0].length;
		double[][] result = new double[2][length];
		for (int j = i - 1; j < length; j++) {
			result[0][j] = data[0][j];
			result[1][j] = (data[1][j] - ma[1][j]) / ma[1][j] * 100;
		}
		result[1][0] = i;

		return result;
	}

	static double[][] RSI(int i) throws IOException {// 569

		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for (int j = i; j < length; j++) {
			result[0][j] = data[0][j];
			double rise = 0;
			double down = 0;
			for (int k = 0; k < i; k++) {
				if (data[1][j - k] - data[1][j - k - 1] > 0) {
					rise += data[1][j - k] - data[1][j - k - 1];
				} else {
					down -= data[1][j - k] - data[1][j - k - 1];
				}
			}
			result[1][j] = (rise / (rise + down)) * 100;
		}
		result[1][0] = i + 1;

		return result;
	}

	static double[][] KDJ(int n, int m1, int m2) throws IOException {// 9,3,3

		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] temp = new double[2][length];
		double[][] result = new double[4][length];
		for (int j = n - 1; j < length; j++) {
			temp[0][j] = data[0][j];
			result[0][j] = data[0][j];
			double top = data[1][j];
			double low = data[1][j];
			for (int k = 0; k < n; k++) {
				if (data[1][j - k] > top) {
					top = data[1][j - k];
				}
				if (data[1][j - k] < low) {
					low = data[1][j - k];
				}
			}
			temp[1][j] = (data[1][j] - low) / (top - low) * 100;// RSV
		}// n-1 begin

		for (int j = n - 1 + m1 - 1; j < length; j++) {
			for (int k = 0; k < m1; k++) {
				result[1][j] += temp[1][j - k];
			}
			result[1][j] /= m1;// K line
		}

		for (int j = n - 1 + m1 - 1 + m2 - 1; j < length; j++) {
			for (int k = 0; k < m2; k++) {
				result[2][j] += result[1][j - k];
			}
			result[2][j] /= m2;// D line
			result[3][j] = 3 * result[1][j] - 2 * result[2][j];// J line
		}

		result[1][0] = n - 1 + m1 - 1 + m2 - 1 + 1;

		return result;
	}

	static double[][] MACD(int shortd, int longd, int midd) throws IOException {// 12,26,9

		double[][] mas = MA(shortd);
		double[][] mal = MA(longd);
		int length = mas[0].length;
		double[][] dif = new double[2][length];
		double[][] dea = new double[2][length];
		double[][] result = new double[2][length];
		for (int j = longd - 1; j < length; j++) {
			dif[0][j] = mas[0][j];
			dif[1][j] = mas[1][j] - mal[1][j];
		}
		for (int j = longd - 1 + midd - 1; j < length; j++) {
			dea[0][j] = dif[0][j];
			for (int k = 0; k < midd; k++) {
				dea[1][j] += dif[1][j - k];
			}
			dea[1][j] /= midd;

			result[0][j] = dea[0][j];
			result[1][j] = dif[1][j] - dea[1][j];
		}

		result[1][0] = longd - 1 + midd - 1 + 1;

		return result;
	}

	/*
	 * KD has been updated to KDJ, so it do not fit. static void
	 * printIndices(double[][] indice,String name){
	 * 
	 * HSSFWorkbook wb = new HSSFWorkbook(); HSSFSheet sheet =
	 * wb.createSheet(name);
	 * 
	 * RowNumber = 0; for(int i = 0;i < indice[0].length;i++){
	 * 
	 * HSSFRow row = sheet.createRow((short)(RowNumber++));//建立新行 Create a row
	 * and put some cells in it. Rows are 0 based. //HSSFCell cell0 =
	 * row.createCell((short)0);//建立新cell
	 * row.createCell((short)0).setCellValue(indice[0][i]);
	 * row.createCell((short)1).setCellValue(indice[1][i]);
	 * 
	 * }
	 * 
	 * try { FileOutputStream fileOut = new
	 * FileOutputStream("file\\"+name+".xls");
	 * wb.write(fileOut);//把Workbook对象输出到文件workbook.xls中 fileOut.close();
	 * 
	 * } catch (Exception e) { // TODO Auto-generated catch block
	 * e.printStackTrace(); } }
	 * 
	 * static void outputIndices() throws IOException{
	 * 
	 * printIndices(MA(5),"dataMA5"); printIndices(BIAS(5),"dataBIAS5");
	 * printIndices(RSI(5),"dataRSI5"); printIndices(KD(5),"dataKD5");
	 * printIndices(MACD(12,26,9),"dataMACD(12,26,9)");
	 * 
	 * }
	 */
	static void printIndices2() throws IOException {

		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("indices");

		double[][] data = pointInsertUsingExcel();
		double[][] ma = MA(5);
		double[][] bias = BIAS(6);
		double[][] rsi = RSI(5);
		double[][] kdj = KDJ(9, 3, 3);
		double[][] macd = MACD(12, 26, 9);

		int begin = (int) ma[1][0];

		if ((int) bias[1][0] > begin) {
			begin = (int) bias[1][0];
		}
		if ((int) rsi[1][0] > begin) {
			begin = (int) rsi[1][0];
		}
		if ((int) kdj[1][0] > begin) {
			begin = (int) kdj[1][0];
		}
		if ((int) macd[1][0] > begin) {
			begin = (int) macd[1][0];
		}

		RowNumber = 0;
		for (int i = begin - 1; i < data[0].length; i++) {

			HSSFRow row = sheet.createRow((short) (RowNumber++));// 建立新行 Create
																	// a row and
																	// put some
																	// cells in
																	// it. Rows
																	// are 0
																	// based.
			// HSSFCell cell0 = row.createCell((short)0);//建立新cell

			row.createCell((short) 0).setCellValue(data[0][i]);
			row.createCell((short) 1).setCellValue(data[1][i]);
			row.createCell((short) 2).setCellValue(ma[1][i]);
			row.createCell((short) 3).setCellValue(bias[1][i]);
			row.createCell((short) 4).setCellValue(rsi[1][i]);
			row.createCell((short) 5).setCellValue(kdj[1][i]);
			row.createCell((short) 6).setCellValue(kdj[2][i]);
			row.createCell((short) 7).setCellValue(kdj[3][i]);
			row.createCell((short) 8).setCellValue(macd[1][i]);

		}

		try {
			FileOutputStream fileOut = new FileOutputStream("file\\"
					+ "indices" + ".xls");
			wb.write(fileOut);// 把Workbook对象输出到文件workbook.xls中
			fileOut.close();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
