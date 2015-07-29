import OOXMLmanager.OOXMLmanager;
import OOXMLmanager.XlsxData;


public class Main {
	private static OOXMLmanager mOOXMLmanager = null;
	public static void main(String argv[]) {
		mOOXMLmanager = new OOXMLmanager();
		
		mOOXMLmanager.setXlsxFile("C:\\UCW_SAMPLE.xlsx");
		
		XlsxData xlsxData = mOOXMLmanager.getXlsxData();				
		
		for (int i = 0; i < xlsxData.getRowSize(); i++) {
			for (int j = 0; j < xlsxData.getColumnSize(); j++) {				
				System.out.printf(" [%3d,%3d] : %20s", i, j, xlsxData.getCellValue(i, j));
			}
			System.out.println("");
		}
		
		xlsxData.deleteEmptyRow();
		System.out.println("delete empty row!");
		for (int i = 0; i < xlsxData.getRowSize(); i++) {
			for (int j = 0; j < xlsxData.getColumnSize(); j++) {
				System.out.printf(" [%3d,%3d] : %20s", i, j, xlsxData.getCellValue(i, j));
			}
			System.out.println("");
		}
		
		xlsxData.deleteEmptyColumn();
		System.out.println("delete empty column!");
		for (int i = 0; i < xlsxData.getRowSize(); i++) {
			for (int j = 0; j < xlsxData.getColumnSize(); j++) {
				System.out.printf(" [%3d,%3d] : %20s", i, j, xlsxData.getCellValue(i, j));
			}
			System.out.println("");
		}
		
		xlsxData.trimData();
		System.out.println("trim data!");
		for (int i = 0; i < xlsxData.getRowSize(); i++) {
			for (int j = 0; j < xlsxData.getColumnSize(); j++) {
				System.out.printf(" [%3d,%3d] : %20s", i, j, xlsxData.getCellValue(i, j));
			}
			System.out.println("");
		}				
	}
}