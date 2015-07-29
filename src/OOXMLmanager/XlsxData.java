package OOXMLmanager;

import java.util.ArrayList;

public class XlsxData {		
	private String[][] mDataArray = null;	

	public XlsxData() {		
		mDataArray = new String[0][0];		
	}

	public void setDataSize(int numOfRow, int numOfColumn) {
		if (numOfRow <= 0 || numOfColumn <= 0) {
			System.out.println("invalid size, row : " + numOfRow + ", column : " + numOfColumn);
		} else {
			mDataArray = new String[numOfRow][numOfColumn];
			for (int i = 0; i < numOfRow; i++) {				
				for (int j = 0; j < numOfColumn; j++) {
					mDataArray[i][j] = "";
				}				
			}		
			System.out.println("SheetSize(" + numOfRow + ", " + numOfColumn + ")");
		}
	}

	public void setCellValue(int row, int column, String value) {
		try {			
			mDataArray[row][column] = value;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public String getCellValue(int row, int column) {
		try {
			return mDataArray[row][column];
		} catch (Exception e) {
			e.printStackTrace();
			return "";
		}
	}

	public int getRowSize() {
		return mDataArray.length;
	}

	public int getColumnSize() {
		return mDataArray[0].length;
	}

	//additional function
	public int deleteEmptyRow() {		
		ArrayList<Integer> notEmptyRowList = new ArrayList<Integer>();

		for (int i = 0; i < mDataArray.length; i++) {				
			for (int j = 0; j < mDataArray[0].length; j++) {
				if (!mDataArray[i][j].equals("")) {
					notEmptyRowList.add(i);
					break;
				}
			}				
		}		

		String[][] newDataArray = new String[notEmptyRowList.size()][mDataArray[0].length];
		for (int i = 0; i < notEmptyRowList.size(); i++) {				
			for (int j = 0; j < mDataArray[0].length; j++) {
				newDataArray[i][j] = mDataArray[notEmptyRowList.get(i)][j];				
			}				
		}

		System.out.println("SheetSize changed! (" + mDataArray.length + ", " + mDataArray[0].length + ") -> (" + newDataArray.length + ", " + newDataArray[0].length + ")");
		
		mDataArray = newDataArray;

		return notEmptyRowList.size();
	}

	public int deleteEmptyColumn() {
		ArrayList<Integer> notEmptyColumnList = new ArrayList<Integer>();

		for (int j = 0; j < mDataArray[0].length; j++) {
			for (int i = 0; i < mDataArray.length; i++) {			
				if (!mDataArray[i][j].equals("")) {
					notEmptyColumnList.add(j);
					break;
				}
			}				
		}		

		String[][] newDataArray = new String[mDataArray.length][notEmptyColumnList.size()];
		for (int j = 0; j < notEmptyColumnList.size(); j++) {
			for (int i = 0; i < mDataArray.length; i++) {							
				newDataArray[i][j] = mDataArray[i][notEmptyColumnList.get(j)];				
			}				
		}

		System.out.println("SheetSize changed! (" + mDataArray.length + ", " + mDataArray[0].length + ") -> (" + newDataArray.length + ", " + newDataArray[0].length + ")");
		
		mDataArray = newDataArray;

		return notEmptyColumnList.size();
	}
	
	public void trimData() {
		for (int j = 0; j < mDataArray[0].length; j++) {
			for (int i = 0; i < mDataArray.length; i++) {			
				mDataArray[i][j] = mDataArray[i][j].trim();
			}				
		}
	}
}

