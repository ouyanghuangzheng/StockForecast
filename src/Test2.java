import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.hssf.usermodel.*;

public class Test2 {


	/**
	 * @param args
	 */
	
	static int RowNumber = 0;
	static int Difference = 0;
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
	
	static void run() throws IOException{
		
		final double limit = 1.5;//you can change the limitation here
		
		//double[][] points = pointInsertUsingTxt();
		double[][] points = pointInsertUsingExcel();
		int j = linearRegression (limit,points);
		if(j != -1){
			double[][] points0 = new double[2][j+1];
			double[][] points1 = new double[2][points[0].length-j];
			
			for(int i = 0;i <= j;i++){
				points0[0][i] = points[0][i];
				points0[1][i] = points[1][i];
			}
			for(int i = j;i <= points[0].length-1;i++){
				points1[0][i-j] = points[0][i];
				points1[1][i-j] = points[1][i];
			}
			
			connection(points[0][0],points[1][0],points[0][j],points[1][j],limit,points0);
			connection(points[0][j],points[1][j],points[0][points[0].length-1],points[1][points[0].length-1],limit,points1);
			output();
		}
	}
	
	static double[][] pointInsertUsingTxt() throws IOException{
		
		BufferedReader input =new BufferedReader(new FileReader("file\\Points.txt"));
		String s, text = new String();
		while((s = input.readLine()) != null)
			text += s;
		input.close();
		String[] location = text.split("\\D");
		double[][] data = new double[2][location.length/2];
		for(int i = 0;i<location.length/2;i++){
			data[0][i] = Double.parseDouble(location[2*i]);
			data[1][i] = Double.parseDouble(location[2*i+1]);
		}
		for(int i = 0;i < data[0].length-1;i++ ){//sorting
			double min = data[0][i];
			int n = i;
			for(int j = i;j <= data[0].length-1;j++){
				if(data[0][j] < min){
					n = j;
				}
			}
			if(n != i){
				double d0 = data[0][i];
				double d1 = data[1][i];
				data[0][i] = data[0][n];
				data[1][i] = data[1][n];
				data[0][n] = d0;
				data[1][n] = d1;
			}
			//System.out.println(data[0][i]+"  "+data[1][i]);
		}
		
		return data;
	}
	
static double[][] pointInsertUsingExcel() throws IOException{
		
		String fileToBeRead = "file\\data.xls";
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeRead)); // 创建对Excel工作簿文件的引用
        HSSFSheet sheet = workbook.getSheet("Sheet2");  // 创建对工作表的引用
        int rows = sheet.getPhysicalNumberOfRows();//获取表格的
        String text = "";
        for (int r = 0; r < rows; r++) {                //循环遍历表格的行
            	HSSFRow row = sheet.getRow(r);         //获取单元格中指定的行对象
            	if (row != null) {
            		int cells = row.getPhysicalNumberOfCells();//获取单元格中指定列对象
            		for (short c = 0; c < cells; c++) {      //循环遍历单元格中的列                  
            			HSSFCell cell = row.getCell((short) c); //获取指定单元格中的列                      
            			if (cell != null&&cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            				text += cell.getNumericCellValue() + ",";
            			}
            		}
            	}
        }
        
        String[] location = text.split(",");
        double[][] data = new double[2][location.length/2];
        for(int i = 0;i < location.length/2;i++){
        	data[0][i] = Double.parseDouble(location[2*i]);
        	data[1][i] = Double.parseDouble(location[2*i+1]);
        	//System.out.println(data[0][i]+"  "+data[1][i]);
    	}
        for(int i = 0;i < data[0].length-1;i++ ){//sorting
        	double min = data[0][i];
        	int n = i;
        	for(int j = i;j <= data[0].length-1;j++){
        		if(data[0][j] < min){
        			n = j;
        		}
        	}
        	if(n != i){
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
	
	static int linearRegression (double limit,double[][] i){
		
		int length = i[0].length;
		double adverage_x = 0;
		double adverage_y = 0;
		double k = 0;
		double b = 0;
		
		for(int j = 0;j < length;j++){
			adverage_x += i[0][j];
		}
		adverage_x /= length;//final
		
		for(int j = 0;j < length;j++){
			adverage_y += i[1][j];
		}
		adverage_y /= length;//final
		
		double k1 = 0;
		for(int j = 0;j < length;j++){
			k1 += (i[0][j]-adverage_x)*(i[1][j]-adverage_y);
		}
		
		double k2 = 0;
		for(int j = 0;j < length;j++){
			k2 += Math.pow((i[0][j]-adverage_x), 2);
		}
		
		k = k1/k2;//final
		
		b = adverage_y-k*adverage_x;//final
		
		System.out.println("The linear regression is "+"y="+k+"x+"+b);
		
		return max(k,b,limit,i);
	}
	
	static int max(double k,double b,double limit,double[][] i){
		double max = 0;
		int number = -1;
		if(i[0].length <= 2){
			return -1;
		}
		for(int j = 0;j < i[0].length;j++){
			if(Math.abs(k*i[0][j]+b-i[1][j]) > max){
				max = Math.abs(k*i[0][j]+b-i[1][j]);
				number = j;
			}
		}
		if(max*Math.pow(1/(1+Math.pow(k, 2)), 1/2) > limit){
			return number;
		}
		return -1;
	}
	
	static void connection(double x1,double y1,double x2,double y2,double limit,double[][] i){
		double k = (y1-y2)/(x1-x2);
		double b = y1-k*x1;
		int j = max(k,b,limit,i);
		if(j != -1){
			double[][] i0 = new double[2][j+1];
			double[][] i1 = new double[2][i[0].length-j];
			
			for(int a = 0;a <= j;a++){
				i0[0][a] = i[0][a];
				i0[1][a] = i[1][a];
			}
			for(int a = j;a <= i[0].length-1;a++){
				i1[0][a-j] = i[0][a];
				i1[1][a-j] = i[1][a];
			}
			
			connection(x1,y1,i[0][j],i[1][j],limit,i0);
			connection(i[0][j],i[1][j],x2,y2,limit,i1);
		}
		else{
			System.out.println("y="+k+"x+"+b+" with points ("+x1+","+y1+") and ("+x2+","+y2+") is fit for limitation!");
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
	
	static void output(){
		
		HSSFWorkbook wb = new HSSFWorkbook();//建立新HSSFWorkbook对象
		HSSFSheet sheet = wb.createSheet("dataoutput");//建立新的sheet对象
		
		/**************************************************************************************/
		HSSFWorkbook wb2 = new HSSFWorkbook();
		HSSFSheet sheet2 = wb2.createSheet("dataoutput2");
		/**************************************************************************************/
		
		while(!list.isEmpty()){
			HSSFRow row = sheet.createRow((short)(RowNumber++));//建立新行 Create a row and put some cells in it. Rows are 0 based.
			//HSSFCell cell0 = row.createCell((short)0);//建立新cell
			double[] group = list.remove(0);
			row.createCell((short)0).setCellValue(group[0]);
			row.createCell((short)1).setCellValue(group[1]);
			row.createCell((short)2).setCellValue(group[2]);
			row.createCell((short)3).setCellValue(group[3]);
			row.createCell((short)4).setCellValue(group[4]);
			row.createCell((short)5).setCellValue(group[5]);
			
			/**************************************************************************************/
			HSSFRow row2 = sheet2.createRow((short)(RowNumber-1+Difference++));
			row2.createCell((short)0).setCellValue(group[2]);
			row2.createCell((short)1).setCellValue(0.5);
			HSSFRow row3 = sheet2.createRow((short)(RowNumber-1+Difference));
			row3.createCell((short)0).setCellValue((group[2]+group[4])/2);
			if(group[0]>=0){
				row3.createCell((short)1).setCellValue(0);
			}
			else{
				row3.createCell((short)1).setCellValue(1);
			}
			/**************************************************************************************/
		}
		
		try {
			FileOutputStream fileOut = new FileOutputStream("file\\dataOutput.xls");
			wb.write(fileOut);//把Workbook对象输出到文件workbook.xls中
			fileOut.close();
			
			/**************************************************************************************/
			FileOutputStream fileOut2 = new FileOutputStream("file\\dataOutput2.xls");
			wb2.write(fileOut2);
			fileOut2.close();
			/**************************************************************************************/
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	
	static double[][] MA(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[0][j] = data[0][j];
			for(int k = 0;k < i;k++){
				result[1][j] += data[1][j-k];
			}
			result[1][j] /= i;
		}
		result[1][0] = i;
		
		return result;
	}
	
	static double[][] BIAS(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		double[][] ma = MA(i);
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[0][j] = data[0][j];
			result[1][j] = (data[1][j]-ma[1][j])/ma[1][j]*100;
		}
		result[1][0] = i;
		
		return result;
	}
	
	static double[][] RSI(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[0][j] = data[0][j];
			double rise = 0;
			double down = 0;
			for(int k = 0;k < i-1;k++){
				if(data[1][j-k]-data[1][j-k-1] > 0){
					rise += data[1][j-k]-data[1][j-k-1];
				}else{
					down -= data[1][j-k]-data[1][j-k-1];
				}
			}
			result[1][j] = (rise/(rise + down))*100;
		}
		result[1][0] = i;
		
		return result;
	}
	
	static double[][] KD(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[0][j] = data[0][j];
			double top = data[1][j];
			double low = data[1][j];
			for(int k = 0;k < i;k++){
				if(data[1][j-k] > top){
					top = data[1][j-k];
				}
				if(data[1][j-k] < low){
					low = data[1][j-k];
				}
			}
			result[1][j] = (data[1][j]-low)/(top-low)*100;
		}
		result[1][0] = i;
		
		return result;
	}
	
	static double[][] MACD(int shortd,int longd,int midd) throws IOException{
		
		double[][] mas = MA(shortd);
		double[][] mal = MA(longd);
		int length = mas[0].length;
		double[][] dif = new double[2][length];
		double[][] dea = new double[2][length];
		double[][] result = new double[2][length];
		for(int j = longd-1;j < length;j++){
			dif[0][j] = mas[0][j];
			dif[1][j] = mas[1][j]-mal[1][j];
		}
		for(int j = longd+midd-2;j < length;j++){
			dea[0][j] = dif[0][j];
			for(int k = 0;k < midd;k++){
				dea[1][j] += dif[1][j-k];
			}
			dea[1][j] /= midd;
			
			result[0][j] = dea[0][j];
			result[1][j] = dif[1][j]-dea[1][j];
		}
		result[1][0] = longd+midd-1;
		
		return result;
	}
	
	static double r(double[][] index) throws IOException{//correlation coefficient
		
		int i = (int)index[1][0]-1;
		double[][] data = pointInsertUsingExcel();
		int length = index[0].length;
		double adverage_x = 0;
		double adverage_y = 0;
		double numerator = 0;
		double denominator1 = 0;
		double denominator2 = 0;
		double r = 0;
		
		for(int j = i;j < length;j++){
			adverage_x += data[1][j];
		}
		adverage_x /= (length-i);//final
		
		for(int j = i;j < length;j++){
			adverage_y += index[1][j];
		}
		adverage_y /= (length-i);//final
		
		for(int j = i;j < length;j++){
			numerator += ((data[1][j]-adverage_x)*(index[1][j]-adverage_y));//final
		}
		
		for(int j = i;j < length;j++){
			denominator1 += Math.pow((data[1][j]-adverage_x),2);
		}
		denominator1 = Math.pow(denominator1,0.5);//final
		
		for(int j = i;j < length;j++){
			denominator2 += Math.pow((index[1][j]-adverage_x),2);
		}
		denominator2 = Math.pow(denominator2,0.5);//final
		
		r = numerator/(denominator1*denominator2);//final
		
		return r;
	}
	
	static void step2and3(double limitation) throws IOException{
		
		//supposing the date is 5
		double[][] ma = MA(5);			double[][] n1 = MA(5);
		double[][] bias = BIAS(5);		double[][] n2 = BIAS(5);
		double[][] rsi = RSI(5);		double[][] n3 = RSI(5);
		double[][] kd = KD(5);			double[][] n4 = KD(5);
		double[][] macd = MACD(5,7,6);	double[][] n5 = MACD(5,7,6);
		
		double rma = Math.pow(r(ma),2);
		double rbias = Math.pow(r(bias),2);
		double rrsi = Math.pow(r(rsi),2);
		double rkd = Math.pow(r(kd),2);
		double rmacd = Math.pow(r(macd),2);//supposing the short date is 5,the long date is 7,and the mid day is 6.
	
		double max = rma;
		double[][] temp;
		
		if(max < rbias){
			max = rbias;
			temp = n1;
			n1 = n2;
			n2 = temp;
		}
		if(max < rrsi){
			max = rrsi;
			temp = n1;
			n1 = n3;
			n3 = temp;
		}
		if(max < rkd){
			max = rkd;
			temp = n1;
			n1 = n4;
			n4 = temp;
		}
		if(max < rmacd){
			max = rmacd;
			temp = n1;
			n1 = n5;
			n5 = temp;
		}
		
		double[] linearRegression = linearRegressionUpgrade(n1);
		double k = linearRegression[0];
		double b = linearRegression[1];
		
		//not finish here
		
	}
	
	static double[] linearRegressionUpgrade (double[][] i){//A linear regression
		
		int a = (int)i[1][0]-1;
		int length = i[0].length;
		double adverage_x = 0;
		double adverage_y = 0;
		double k = 0;
		double b = 0;
		
		for(int j = a;j < length;j++){
			adverage_x += i[0][j];
		}
		adverage_x /= length-a;//final
		
		for(int j = a;j < length;j++){
			adverage_y += i[1][j];
		}
		adverage_y /= length-a;//final
		
		double k1 = 0;
		for(int j = a;j < length;j++){
			k1 += (i[0][j]-adverage_x)*(i[1][j]-adverage_y);
		}
		
		double k2 = 0;
		for(int j = a;j < length;j++){
			k2 += Math.pow((i[0][j]-adverage_x), 2);
		}
		
		k = k1/k2;//final
		
		b = adverage_y-k*adverage_x;//final
		
		System.out.println("The linear regression is "+"y="+k+"x+"+b);
		
		double[] result = {k,b};
		
		return result;
	}
	
	static void linearRegressionUpgrade2 (double[][] i){//Binary linear regression
		
	}
	
	static double f_value(int k,int b,double[][] i){
		
		int a = (int)i[1][0]-1;
		int length = i[0].length;
		double adverage_y = 0;
		double SSR = 0;
		double SSE = 0;
		double F_value = 0;
		
		for(int j = a;j < length;j++){
			adverage_y += i[1][j];
		}
		adverage_y /= (length-a);//final
		
		for(int j = a;j < length;j++){
			SSR += Math.pow((k*i[0][j]+b-adverage_y), 2);//final
		}
		
		for(int j = a;j < length;j++){
			SSE += Math.pow((k*i[0][j]+b-i[1][j]), 2);//final
		}
		
		F_value = (SSR/1)/(SSE/(length-a-2));
		
		return F_value;
	}
}
