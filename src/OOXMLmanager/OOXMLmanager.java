package OOXMLmanager;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class OOXMLmanager {
	private final String SHARED_STRING_PATH = "/xl/sharedStrings.xml";	
	private final String TARGET_SHEET_PATH = "/xl/worksheets/";
	private final String TEMP_ZIP_FILE_NAME = "tempXlsxFile.zip";

	private ArrayList<String> mSharedStringList = null;
	private String mRootPath = "";
	private String mXlsxFilePath = "";
	private String mTempDirPath = "";
	private String mUnZipDirPath = "";
	private String mOpenSheetName = "sheet1.xml";
	private XlsxData mXlsxData = null;	

	public OOXMLmanager() {
		mSharedStringList = new ArrayList<String>();			
		mXlsxData = new XlsxData();
	}

	/**
	 * 데이터를 추출할 xlsx 파일을 세팅해준다. 
	 * @param path xlsx 파일이 존재하는 경로
	 */
	public void setXlsxFile(String path) {
		setXlsxFile(new File(path));
	}

	/**
	 * 데이터를 추출할 xlsx 파일을 세팅해준다.
	 * @param file xlsx 파일 객체
	 */
	public void setXlsxFile(File file) {		
		try {
			mRootPath = file.getParent();
			mXlsxFilePath = file.getAbsolutePath();
			mTempDirPath = mRootPath + "/tempXlsxFile";
			mUnZipDirPath = mTempDirPath + "/tempXlsxFile";

			deleteDirectory(mTempDirPath);
			makeDirectory(mTempDirPath);

			fileCopy(mXlsxFilePath, mTempDirPath + "/" + TEMP_ZIP_FILE_NAME);		
			unZip(mTempDirPath + "/" + TEMP_ZIP_FILE_NAME, mUnZipDirPath);

			readSharedStrings();
			readSheetData();

			deleteDirectory(mTempDirPath);
		} catch (Exception exception) {
			exception.printStackTrace();
		}
	}

	/**
	 * 데이터를 추출할 sheet의 번호를 설정한다.(1부터 시작)
	 * 사용자가 저장해논 sheet명은 반영 되지않고 sheet1, sheet2 이런 식으로 저장되기 때문.
	 * @param sheetName 엑셀 시트명
	 */
	public void setOpenSheetName(String sheetNum) {
		int iSheetNum = 1;
		if (sheetNum != null && sheetNum.length() > 0) {
			iSheetNum = Integer.parseInt(sheetNum);
		}
		mOpenSheetName = "sheet" + iSheetNum;
	}

	/**
	 * xlsx 파일에서 추출한 데이터를 반환.
	 * @return XlsxData 형식으로 이루어진 데이터.
	 */
	public XlsxData getXlsxData() {
		return mXlsxData;
	}

	/**
	 * 디렉토리 생성
	 * @param DirectoryPath 생성할 디렉토리 경로
	 */
	private void makeDirectory(String DirectoryPath) {
		File directory = new File(DirectoryPath);

		//create directory if not exist
		if (!directory.isDirectory()) {
			directory.mkdirs();
		}
	}

	/**
	 * 디렉토리 삭제
	 * @param DirectoryPath 삭제할 디렉토리 경로
	 * @return 삭제 성공시 true, 실패시 false
	 */
	private boolean deleteDirectory(String DirectoryPath) {
		File directory = new File(DirectoryPath);

		if(!directory.exists()) {
			return false;
		}

		File[] files = directory.listFiles();
		for (File file : files) {
			if (file.isDirectory()) {
				deleteDirectory(file.getAbsolutePath());
			} else {
				file.delete();
			}
		}

		return directory.delete();
	}	

	/**
	 * 파일 복사 
	 * @param inFileName 복사활 파일 경로
	 * @param outFileName 복사될 파일 경로
	 */
	private void fileCopy(String inFileName, String outFileName) {
		try {
			FileInputStream fis = new FileInputStream(inFileName);
			FileOutputStream fos = new FileOutputStream(outFileName);

			int data = 0;
			while ((data=fis.read()) != -1) {
				fos.write(data);
			}
			fis.close();
			fos.close();

		} catch (IOException e) { 
			e.printStackTrace();
		}
	}

	/**
	 * sharedStrings.xml 안의 스트링 데이터를 읽어와 리스트에 세팅
	 */
	private void readSharedStrings() {
		try {
			File file = new File(mUnZipDirPath + SHARED_STRING_PATH);

			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
			DocumentBuilder db = dbf.newDocumentBuilder();
			Document doc = db.parse(file);

			doc.getDocumentElement().normalize();

			NodeList siList = doc.getElementsByTagName("si");			

			for (int i = 0; i < siList.getLength(); i++) {
				Node siNode = siList.item(i);
				Element siElmnt = (Element) siNode;

				NodeList tList = siElmnt.getElementsByTagName("t");
				Node tNode = tList.item(0);
				Element tElmnt = (Element) tNode;

				NodeList leafList = tElmnt.getChildNodes();
				String leafValue = ((Node) leafList.item(0)).getNodeValue();
				mSharedStringList.add(leafValue.trim());				
			}					

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * sheet1.xml 안의 데이터를 읽어온다.(shardStrings.xml 에서 가져온 string 값들과의 매칭작업까지)
	 */
	private void readSheetData() {
		try {
			File file = new File(mUnZipDirPath + TARGET_SHEET_PATH + mOpenSheetName);

			DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
			DocumentBuilder db = dbf.newDocumentBuilder();
			Document doc = db.parse(file);

			doc.getDocumentElement().normalize();

			NodeList sheetDataList = doc.getElementsByTagName("sheetData");
			Node sheetDataNode = sheetDataList.item(0);
			Element sheetDataElmnt = (Element) sheetDataNode;

			NodeList rowList = sheetDataElmnt.getElementsByTagName("row");
			//set size of array[][]			
			String dimension = ((Element) doc.getElementsByTagName("dimension").item(0)).getAttribute("ref");
			if (dimension != null && dimension.matches("^[a-zA-Z]+[0-9]+:[a-zA-Z]+[0-9]+$")) {
				//dimension ref 정보가 있을경우
				int[] sizeInfo = getSheetSize(((Element) doc.getElementsByTagName("dimension").item(0)).getAttribute("ref"));			
				mXlsxData.setDataSize(sizeInfo[0], sizeInfo[1]);
			} else {
				//dimension ref 정보가 없을경우
				int maxColumnCount = 0;
				for (int i = 0; i < rowList.getLength(); i++) {
					Node rowNode = rowList.item(i);				
					Element rowElmnt = (Element) rowNode;

					NodeList cellList = rowElmnt.getElementsByTagName("c");
					maxColumnCount = cellList.getLength() > maxColumnCount ? cellList.getLength() : maxColumnCount; 
				}			
				mXlsxData.setDataSize(rowList.getLength(), maxColumnCount);			
			}

			//extract data
			for (int i = 0; i < rowList.getLength(); i++) {
				Node rowNode = rowList.item(i);				
				Element rowElmnt = (Element) rowNode;

				NodeList cellList = rowElmnt.getElementsByTagName("c");
				for (int j = 0; j < cellList.getLength(); j++) {
					Node cellNode = cellList.item(j);
					Element cellElmnt = (Element) cellNode;

					NodeList valueList = cellElmnt.getElementsByTagName("v");
					Node valueNode = valueList.item(0);
					Element valueElmnt = (Element) valueNode;

					if (valueElmnt != null) {
						NodeList leafList = valueElmnt.getChildNodes();
						String leafValue = ((Node) leafList.item(0)).getNodeValue();											

						//extract cell value if type is string
						if (cellElmnt.getAttribute("t").equalsIgnoreCase("s")) {
							leafValue = mSharedStringList.get(Integer.parseInt(leafValue));
							//System.out.printf(" [%3d,%3d] : %20s", (i + 1), (j + 1), leafValue);						
						} else {
							//System.out.printf(" [%3d,%3d] : %20s", (i + 1), (j + 1), leafValue);
						}					

						mXlsxData.setCellValue(i, j, leafValue);
					} else {
						//System.out.printf(" [%3d,%3d] : null!", (i + 1), (j + 1));
					}
				}				
				//System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * zip으로 변환된 xlsx 파일을 압축해제 한다.
	 * @param zipFilePath zip 파일경로
	 * @param outputDirPath 압축을 풀 경로
	 */
	private void unZip(String zipFilePath, String outputDirPath) {
		try {
			System.out.println("Begin unzip "+ zipFilePath + " into "+outputDirPath);

			ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFilePath));
			ZipEntry ze = zis.getNextEntry();

			while (ze != null){
				String entryName = ze.getName();

				System.out.print("Extracting " + entryName + " -> " + outputDirPath + File.separator +  entryName + "...");

				File f = new File(outputDirPath + File.separator +  entryName);
				//create all folder needed to store in correct relative path.
				f.getParentFile().mkdirs();
				FileOutputStream fos = new FileOutputStream(f);
				int len;
				byte buffer[] = new byte[1024];
				while ((len = zis.read(buffer)) > 0) {
					fos.write(buffer, 0, len);
				}
				fos.close();   

				System.out.println("OK!");

				ze = zis.getNextEntry();
			}
			zis.closeEntry();
			zis.close();

			System.out.println( zipFilePath + " unzipped successfully");
		} catch (Exception exception) {
			exception.printStackTrace();
		}
	}	
	 
	/**
	 * dimension 태그의 ref속성 값을 가지고 sheet의 row, column 수를 반환한다.
	 * @param dimension dimension 태그내의 ref 속성값
	 * @return rowCount, columnCount
	 */
	private int[] getSheetSize(String dimension) {
		int[] sizeInfo = new int[2];		
		
		String[] rangeInfo = dimension.split(":");
		int startRow = Integer.parseInt(rangeInfo[0].replaceAll("[^0-9]", ""));
		int endRow = Integer.parseInt(rangeInfo[1].replaceAll("[^0-9]", ""));
		int startColumn = columnNameToNumber(rangeInfo[0].replaceAll("[0-9]", ""));
		int endColumn = columnNameToNumber(rangeInfo[1].replaceAll("[0-9]", ""));
		
		sizeInfo[0] = endRow - startRow + 1;
		sizeInfo[1] = endColumn - startColumn + 1;
		
		return sizeInfo;
	}
	   
	/**
	 * 시트의 알파벳 컬럼명을 숫자값으로 바꿔준다. A=1, B=2, ...
	 * @param colName 컬럼명(알파벳)
	 * @return 변환된 숫자값
	 */
    public int columnNameToNumber(String colName)  
    {  
      int result = 0;  

      for (int i = 0; i < colName.length(); i++)  
      {  
        result *= 26;  
        char letter = colName.charAt(i);  
        // See if it's out of bounds.  
        if (letter < 'A') {
        	letter = 'A';  
        }
        if (letter > 'Z') {
        	letter = 'Z';  
        }
        // Add in the value of this letter.  
        result += (int) letter - (int) 'A' + 1;  
      }  
      return result;  
    }  
}
