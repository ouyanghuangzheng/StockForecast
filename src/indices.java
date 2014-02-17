import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class indices {

	static double[][] MA(int i) throws IOException{
	
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[j][0] = data[j][0];
			for(int k = 0;k < i;k++){
				result[j][1] += data[j-k][1];
			}
			result[j][1] /= i;
		}
		return result;
	}
	
	static double[][] BIAS(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		double[][] ma = MA(i);
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[j][0] = data[j][0];
			result[j][1] = (data[j][1]-ma[j][1])/ma[j][1]*100;
		}
		return result;
	}
	
	static double[][] RSI(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[j][0] = data[j][0];
			double rise = 0;
			double down = 0;
			for(int k = 0;k <= i-2;k++){
				if(data[j-k][1]-data[j-k-1][1] > 0){
					rise += data[j-k][1]-data[j-k-1][1];
				}else{
					down -= data[j-k][1]-data[j-k-1][1];
				}
			}
			result[j][1] = (rise/(rise + down))*100;
		}
		return result;
	}
	
	static double[][] KD(int i) throws IOException{
		
		double[][] data = pointInsertUsingExcel();
		int length = data[0].length;
		double[][] result = new double[2][length];
		for(int j = i-1;j < length;j++){
			result[j][0] = data[j][0];
			double top = data[j][1];
			double low = data[j][1];
			for(int k = 0;k <= i-1;k++){
				if(data[j-k][1] > top){
					top = data[j-k][1];
				}
				if(data[j-k][1] < low){
					low = data[j-k][1];
				}
			}
			result[j][1] = (data[j][1]-low)/(top-low)*100;
		}
		return result;
	}
	
	static double[][] MACD(int shortd,int longd,int midd) throws IOException{
		
		double[][] mas = MA(shortd);
		double[][] mal = MA(longd);
		int length = mas[0].length;
		double[][] dif = new double[2][length];
		double[][] dea = new double[2][length];
		double[][] macd = new double[2][length];
		for(int j = 0;j < length;j++){
			dif[j][0] = mas[j][0];
			dif[j][1] = mas[j][1]-mal[j][1];
		}
		for(int j = longd+midd-2;j < length;j++){
			dea[j][0] = dif[j][0];
			for(int k = 0;k < midd;k++){
				dea[j][1] += dif[j-k][1];
			}
			dea[j][1] /= midd;
		}
		for(int k = longd+midd-2;k < length;k++){
			macd[k][0] = dea[k][0];
			macd[k][1] = dif[k][1]-dea[k][1];
		}
		return macd;
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
}
