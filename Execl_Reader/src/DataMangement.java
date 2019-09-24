
public class DataMangement {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Xls_Reader xls = new Xls_Reader("E:\\Shilpa\\Selenium\\Selenium Notes\\Selenium example codes\\Excel_Reader\\data.xlsx");
		String sheetName = "Data";
		String testCaseName = "TestC";
		
		int testStartRowNum = 1;
		String x = xls.getCellData(sheetName, 0, 2);
		System.out.println(x);
		
		while(!xls.getCellData(sheetName, 0, testStartRowNum).equals(testCaseName)) {
			testStartRowNum++;
		}
		System.out.println("Test Case start from Row"+testStartRowNum);
		int colStartRowNum = testStartRowNum + 1;
		int dataStartRowNum = testStartRowNum + 2;
		
		//calculate rows of data
		
		int rows = 1;
		
		while(!xls.getCellData(sheetName, 0, dataStartRowNum+rows).equals("")) {
			rows ++;
		}
		System.out.println("Total rows are -"+rows);
		
		int cols = 0;
		//calculate col of 
		while(!xls.getCellData(sheetName, cols, colStartRowNum).equals("")) {
			cols ++;
		}
		System.out.println("Total cols are -"+cols);
		
		//read the data 
		
		for(int rNum = dataStartRowNum; rNum < dataStartRowNum+rows; rNum++) {
			for(int cNum = 0; cNum < cols; cNum++) {
				System.out.println(xls.getCellData(sheetName, cNum, rNum));
			}
		}
	}
}
