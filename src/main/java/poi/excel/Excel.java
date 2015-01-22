package poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

public class Excel {

	public static void main(String[] args) throws Exception {
		String fileName = "驾驶人台账.xls";
		copyExcel(fileName);
	}

	
	public static void copyExcel(String fileName) throws Exception{
		System.out.println("+++++++++++++++++++++++++++++++++++++++++++++++++");
		System.out.print("excel " + fileName + ">>>>>>>>>>");
		FileInputStream fileInputStream = new FileInputStream(fileName);
		HSSFWorkbook sourceWorkbook = new HSSFWorkbook(fileInputStream);
		int index = fileName.indexOf("台账.xls");
		if(index == -1){
			System.out.println("[error]:filename error " + fileName);
			return;
		}
		String newFileName = fileName.substring(0, index) + "信息表.xls";
		System.out.println(newFileName);
		
		FileInputStream templateFileInputStream = new FileInputStream("template.xls");
		HSSFWorkbook templateSourceWorkbook = new HSSFWorkbook(templateFileInputStream);
		HSSFSheet templateSheet = templateSourceWorkbook.getSheetAt(0);
		
		Map<Integer, HSSFCellStyle> styleMap = new HashMap<Integer, HSSFCellStyle>();
		
		HSSFWorkbook targetWorkbook = new HSSFWorkbook();
		
		int numberOfSheets = sourceWorkbook.getNumberOfSheets();
		int numberOfHandled = 0;
		HSSFCellStyle hssfCellStyle = targetWorkbook.createCellStyle();
		for(int i = 0; i < numberOfSheets; i++){
			HSSFSheet sourceSheet = sourceWorkbook.getSheetAt(i);
			HSSFSheet targetSheet = targetWorkbook.createSheet(sourceSheet.getSheetName());
			copySheet(sourceSheet, targetSheet,hssfCellStyle, templateSheet, styleMap);
				
			
			numberOfHandled++;
		}
		
		System.out.print("总数:" + numberOfSheets + " 已处理:" + numberOfHandled);
		if(numberOfSheets != numberOfHandled){
			System.out.println(" xxxxxxx");
		}
		
		FileOutputStream out = new FileOutputStream(new File(newFileName));
		targetWorkbook.write(out);
		out.close();
	}
	
	public static void copySheet(HSSFSheet sourceSheet, HSSFSheet targetSheet, HSSFCellStyle hssfCellStyle,
			HSSFSheet templateshSheet, Map<Integer, HSSFCellStyle> styleMap) throws Exception{
		System.out.println("sheet " + sourceSheet.getSheetName());
		Iterator<Row> rowiIterator = sourceSheet.iterator();
		while(rowiIterator.hasNext()){
			Row row = rowiIterator.next();
			Cell cell = row.getCell(row.getFirstCellNum());
			if(cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().equals("序号"))
				break;
		}
		for(int rowNumber = 1;rowiIterator.hasNext();rowNumber++){
			System.out.print("row " + rowNumber + "-");
			Row row = rowiIterator.next();
			if(row == null)
				continue;
			int startRow = (rowNumber-1)*16 + 1;
			
			
			if(row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK){
				System.out.println("有空行");
				continue;
			}
			copySubSheets(targetSheet, templateshSheet, startRow, styleMap);
			
			System.out.println(row.getCell(0).getNumericCellValue());
			
			Cell namecCell = row.getCell(2);
			Cell carTypeCell = row.getCell(3);
			Cell idCell = row.getCell(4);
			Cell addrCell = row.getCell(5);
			Cell telCell = row.getCell(6);
			
			Cell firstTimeCell = row.getCell(14);
			Cell expiredTimeCell = row.getCell(12);
			Cell fileIDCell = row.getCell(13);
			
			
			String id = getCellValue(idCell);
			String sex = null;
			try {
				sex = id.substring(16, 17);
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			if(Integer.parseInt(sex) % 2 == 0)
				sex = "女";
			else {
				sex = "男";
			}
			String birth = id.substring(6, 14);
			
			DateFormat format = new SimpleDateFormat("yyyy-MM-dd");
			
			Row targetRow1 = targetSheet.getRow(startRow);
			Cell targetNameCell = targetRow1.getCell(2);
			targetNameCell.setCellValue(getCellValue(namecCell));
			Cell targetSexCell = targetRow1.getCell(5);
			targetSexCell.setCellValue(sex);
			Cell targetBirthCell = targetRow1.getCell(11);
			targetBirthCell.setCellValue(birth);
			
			Row targetRow2 = targetSheet.getRow(startRow + 1);
			Cell targetIDCell = targetRow2.getCell(5);
			targetIDCell.setCellValue(id);
			
			Row targetRow3 = targetSheet.getRow(startRow + 2);
			Cell targetAddrCell = targetRow3.getCell(2);
			targetAddrCell.setCellValue(getCellValue(addrCell));
			Cell targetFirstTimeCell = targetRow3.getCell(19);
			if(firstTimeCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
				targetFirstTimeCell.setCellValue(format.format(firstTimeCell.getDateCellValue()));
			}else {
				targetFirstTimeCell.setCellValue(getCellValue(firstTimeCell));
			}
			
			Row targetRow4 = targetSheet.getRow(startRow + 3);
			Cell targetExpiredTimeCell = targetRow4.getCell(2);
			targetExpiredTimeCell.setCellValue(format.format(expiredTimeCell.getDateCellValue()));
			Cell targetTelCell = targetRow4.getCell(19);
			if(telCell.getCellType() == Cell.CELL_TYPE_NUMERIC){
				BigDecimal bigDecimal = new BigDecimal(telCell.getNumericCellValue());
				targetTelCell.setCellValue(bigDecimal.toString());
			}else {
				targetTelCell.setCellValue(telCell.getStringCellValue());
			}
			
			Row targetRow5 = targetSheet.getRow(startRow + 4);
			Cell targetFileIDCell = targetRow5.getCell(2);
			targetFileIDCell.setCellValue(getCellValue(fileIDCell));
			Cell targetCarTypeCell = targetRow5.getCell(11);
			targetCarTypeCell.setCellValue(getCellValue(carTypeCell));
			Cell targetCheckExpiredTimeCell = targetRow5.getCell(19);
			targetCheckExpiredTimeCell.setCellValue(format.format(expiredTimeCell.getDateCellValue()));
			
		}
	}
	
	
	public static String getCellValue(Cell cell){
		String result = "";
		if(cell == null)
			return result;
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			result = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			result = String.valueOf(cell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			break;
		default:
			break;
			}
		return result;
	}
	
	
	
	/**
	 * @param newSheet
	 *            the sheet to create from the copy.
	 * @param sheet
	 *            the sheet to copy.
	 */
	public static void copySubSheets(HSSFSheet newSheet, HSSFSheet sheet,
			int startRow, Map<Integer, HSSFCellStyle> styleMap) {
		copySubSheets(newSheet, sheet, true, startRow, styleMap);
	}

	/**
	 * @param newSheet
	 *            the sheet to create from the copy.
	 * @param sheet
	 *            the sheet to copy.
	 * @param copyStyle
	 *            true copy the style.
	 */
	public static void copySubSheets(HSSFSheet newSheet, HSSFSheet sheet,
			boolean copyStyle, int startRow, Map<Integer, HSSFCellStyle> styleMap) {
		int maxColumnNum = 0;
//		Map<Integer, HSSFCellStyle> styleMap = (copyStyle) ? new HashMap<Integer, HSSFCellStyle>()
//				: null;
		for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
			HSSFRow srcRow = sheet.getRow(i);
			HSSFRow destRow = newSheet.createRow(i + startRow - 1);
			if (srcRow != null) {
				copyRow(sheet, newSheet, srcRow, destRow, styleMap);
				if (srcRow.getLastCellNum() > maxColumnNum) {
					maxColumnNum = srcRow.getLastCellNum();
				}
			}
		}
		for (int i = 0; i <= maxColumnNum; i++) {
			newSheet.setColumnWidth(i, sheet.getColumnWidth(i));
		}
	}

	/**
	 * @param srcSheet
	 *            the sheet to copy.
	 * @param destSheet
	 *            the sheet to create.
	 * @param srcRow
	 *            the row to copy.
	 * @param destRow
	 *            the row to create.
	 * @param styleMap
	 *            -
	 */
	public static void copyRow(HSSFSheet srcSheet, HSSFSheet destSheet,
			HSSFRow srcRow, HSSFRow destRow,
			Map<Integer, HSSFCellStyle> styleMap) {
		// manage a list of merged zone in order to not insert two times a
		// merged zone
		Set<CellRangeAddressWrapper> mergedRegions = new TreeSet<CellRangeAddressWrapper>();
		destRow.setHeight(srcRow.getHeight());
		// reckoning delta rows
		int deltaRows = destRow.getRowNum() - srcRow.getRowNum();
		// pour chaque row
		for (int j = srcRow.getFirstCellNum(); j <= srcRow.getLastCellNum(); j++) {
			HSSFCell oldCell = srcRow.getCell(j); // ancienne cell
			HSSFCell newCell = destRow.getCell(j); // new cell
			if (oldCell != null) {
				if (newCell == null) {
					newCell = destRow.createCell(j);
				}
				// copy chaque cell
				copyCell(oldCell, newCell, styleMap);
				// copy les informations de fusion entre les cellules
				// System.out.println("row num: " + srcRow.getRowNum() +
				// " , col: " + (short)oldCell.getColumnIndex());
				CellRangeAddress mergedRegion = getMergedRegion(srcSheet,
						srcRow.getRowNum(), (short) oldCell.getColumnIndex());

				if (mergedRegion != null) {
					// System.out.println("Selected merged region: " +
					// mergedRegion.toString());
					CellRangeAddress newMergedRegion = new CellRangeAddress(
							mergedRegion.getFirstRow() + deltaRows,
							mergedRegion.getLastRow() + deltaRows,
							mergedRegion.getFirstColumn(),
							mergedRegion.getLastColumn());
					// System.out.println("New merged region: " +
					// newMergedRegion.toString());
					CellRangeAddressWrapper wrapper = new CellRangeAddressWrapper(
							newMergedRegion);
					if (isNewMergedRegion(wrapper, mergedRegions)) {
						mergedRegions.add(wrapper);
						destSheet.addMergedRegion(wrapper.range);
					}
				}
			}
		}
	}

	/**
	 * @param oldCell
	 * @param newCell
	 * @param styleMap
	 */
	public static void copyCell(HSSFCell oldCell, HSSFCell newCell,
			Map<Integer, HSSFCellStyle> styleMap) {
		if (styleMap != null) {
			if (oldCell.getSheet().getWorkbook() == newCell.getSheet()
					.getWorkbook()) {
				newCell.setCellStyle(oldCell.getCellStyle());
			} else {
				int stHashCode = oldCell.getCellStyle().hashCode();
				HSSFCellStyle newCellStyle = styleMap.get(stHashCode);
				if (newCellStyle == null) {
					newCellStyle = newCell.getSheet().getWorkbook()
							.createCellStyle();
					newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
					styleMap.put(stHashCode, newCellStyle);
				}
				newCell.setCellStyle(newCellStyle);
			}
		}
		switch (oldCell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			newCell.setCellValue(oldCell.getStringCellValue());
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			newCell.setCellType(HSSFCell.CELL_TYPE_BLANK);
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case HSSFCell.CELL_TYPE_FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());
			break;
		default:
			break;
		}

	}

	/**
	 * Récupère les informations de fusion des cellules dans la sheet source
	 * pour les appliquer à la sheet destination... Récupère toutes les zones
	 * merged dans la sheet source et regarde pour chacune d'elle si elle se
	 * trouve dans la current row que nous traitons. Si oui, retourne l'objet
	 * CellRangeAddress.
	 * 
	 * @param sheet
	 *            the sheet containing the data.
	 * @param rowNum
	 *            the num of the row to copy.
	 * @param cellNum
	 *            the num of the cell to copy.
	 * @return the CellRangeAddress created.
	 */
	public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum,
			short cellNum) {
		for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
			CellRangeAddress merged = sheet.getMergedRegion(i);
			if (merged.isInRange(rowNum, cellNum)) {
				return merged;
			}
		}
		return null;
	}

	/**
	 * Check that the merged region has been created in the destination sheet.
	 * 
	 * @param newMergedRegion
	 *            the merged region to copy or not in the destination sheet.
	 * @param mergedRegions
	 *            the list containing all the merged region.
	 * @return true if the merged region is already in the list or not.
	 */
	private static boolean isNewMergedRegion(
			CellRangeAddressWrapper newMergedRegion,
			Set<CellRangeAddressWrapper> mergedRegions) {
		return !mergedRegions.contains(newMergedRegion);
	}

}

class CellRangeAddressWrapper implements Comparable<CellRangeAddressWrapper> {

	public CellRangeAddress range;

	/**
	 * @param theRange
	 *            the CellRangeAddress object to wrap.
	 */
	public CellRangeAddressWrapper(CellRangeAddress theRange) {
		this.range = theRange;
	}

	/**
	 * @param o
	 *            the object to compare.
	 * @return -1 the current instance is prior to the object in parameter, 0:
	 *         equal, 1: after...
	 */
	public int compareTo(CellRangeAddressWrapper o) {

		if (range.getFirstColumn() < o.range.getFirstColumn()
				&& range.getFirstRow() < o.range.getFirstRow()) {
			return -1;
		} else if (range.getFirstColumn() == o.range.getFirstColumn()
				&& range.getFirstRow() == o.range.getFirstRow()) {
			return 0;
		} else {
			return 1;
		}

	}

}