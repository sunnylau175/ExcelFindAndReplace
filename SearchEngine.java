import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JRadioButton;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.table.DefaultTableModel;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.RecordFormatException;
import org.apache.poi.hssf.record.cf.CellRangeUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class SearchEngine implements Runnable {
	
	private ArrayList<ExcelFileInfo> fileInfo;	//temp variable
	private ArrayList<ExcelFileInfo> listFileInfo;
	private ArrayList<ArrayList<Dependencies>> relatedFormulaCells;
	private ArrayList<Dependencies> dependenciesHistory;
	private JTable tableResult;
	private DefaultTableModel tableResultModel;
	private JButton buttonFind;
	private JButton buttonReplace;
	private JButton buttonReplaceAll;
	private JButton buttonBrowse;
	private boolean hasResult = false;
	private JLabel lblProgress;
	public static final int ACTION_FIND = 0;
	public static final int ACTION_REPLACE = 1;
	public static final int ACTION_REPLACE_ALL = 2;
	public static final int ACTION_UPDATE_RELATIVE_FORMULA = 3;
	public static final int ACTION_UPDATE_ALL_FORMULA = 4;
	private int requestedAction = ACTION_FIND;
	private int requestedFormulaAction = ACTION_UPDATE_RELATIVE_FORMULA;
	private String location, find, replace;
	private boolean includeSubdirectories = false, caseSensitive = false, exactCellMatch = false;
	private volatile boolean terminateThread = false;
	private JButton btnStop;
	private JCheckBox chckbxIncludeSubdirectories;
	private JCheckBox chckbxCaseSensitive;
	private JCheckBox chckbxExactCellMatch;
	private JRadioButton rdbtnUpdateRelativeFormula;
	private JRadioButton rdbtnUpdateAllFormula;
	

	public SearchEngine(JTable table, JButton find, JButton replace, JButton replaceAll, JButton browse, JLabel progress, JButton stop, JCheckBox includeSubdirectories, JCheckBox caseSensitive, JCheckBox exactCellMatch, JRadioButton updateRelativeFormula, JRadioButton updateAllFormula) {
		fileInfo = new ArrayList<ExcelFileInfo>();
		listFileInfo = new ArrayList<ExcelFileInfo>();
		relatedFormulaCells = new ArrayList<ArrayList<Dependencies>>();
		dependenciesHistory = new ArrayList<Dependencies>();
		
		tableResult = table;
		tableResultModel = (DefaultTableModel) table.getModel();
		buttonFind = find;
		buttonReplace = replace;
		buttonReplaceAll = replaceAll;
		buttonBrowse = browse;
		lblProgress = progress;
		btnStop = stop;
		chckbxIncludeSubdirectories = includeSubdirectories;
		chckbxCaseSensitive = caseSensitive;
		chckbxExactCellMatch = exactCellMatch;
		
		rdbtnUpdateRelativeFormula = updateRelativeFormula;
		rdbtnUpdateAllFormula = updateAllFormula;
		
	}
	
	public void terminateSearchEngine(boolean terminate) {
		terminateThread = terminate;
		
		lblProgress.setText("Stopping...");
		
		
	}
	
	public void run() {
		
		String fileName, fileLocation;
		String sheet, cell;
		String temp;
		
		disableAllButtons();
		
		switch (requestedAction) {
		
			case ACTION_FIND:
				lblProgress.setText("Searching Text...");
				lblProgress.setVisible(true);
				btnStop.setVisible(true);
				
				clearSearchResult();
				if (!find.equals(""))
					listSearchResult(new File(location), find, includeSubdirectories, caseSensitive, exactCellMatch);
				
				//search dependencies
				lblProgress.setText("Searching Dependencies...");
				
				for (ExcelFileInfo f : listFileInfo) {
					File excelFile = new File(f.location, f.fileName);
					try {
						dependenciesHistory.clear();
						searchDependencies(f.sheet, f.cell, excelFile, f.dependentCell);
						
					} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					
				}
				
				if (terminateThread) {
					//System.out.println("Thread interrupting...");
					lblProgress.setVisible(false);
					btnStop.setVisible(false);
					enableAllButtons();
					terminateThread = false;
					return;
				}
				
				
			
				break;
		
			case ACTION_REPLACE:
				lblProgress.setText("Replacing...");
				lblProgress.setVisible(true);
				btnStop.setVisible(true);
				
				// read dependencies from tree into 2d array list
				for (ExcelFileInfo f : listFileInfo) {
					ArrayList<Dependencies> saveDependencies = new ArrayList<Dependencies>();
					readDependencies(f.dependentCell, saveDependencies);
					relatedFormulaCells.add(saveDependencies);
					
					
				}
				
				
				
				int[] selectedRows = tableResult.getSelectedRows();
			
				for (int i : selectedRows) {
					fileName = tableResult.getValueAt(i, 0).toString();
					fileLocation = tableResult.getValueAt(i, 1).toString();
					temp = tableResult.getValueAt(i, 2).toString();
					sheet = temp.substring(0, temp.indexOf("!"));
					cell = temp.substring(temp.indexOf("!"), temp.length());
				
					try {
						File excelFile = new File(fileLocation, fileName);
						replaceString(excelFile, find, replace, sheet, cell, caseSensitive);
						
						if (requestedFormulaAction == ACTION_UPDATE_RELATIVE_FORMULA)
							recalculateFormulas(excelFile, i);
						
						if (terminateThread) {
							//System.out.println("Thread interrupting...");
							lblProgress.setVisible(false);
							btnStop.setVisible(false);
							enableAllButtons();
							terminateThread = false;
							return;
						}
						
						Object replaced = "Yes";
						tableResultModel.setValueAt(replaced, i, 3);
					} catch (InvalidFormatException | IOException e) {
						// TODO Auto-generated catch block
						//e.printStackTrace();
						Object replaced = "No";
						tableResultModel.setValueAt(replaced, i, 3);
					}
				}
			
				break;
			
		
			case ACTION_REPLACE_ALL:
				lblProgress.setText("Replacing...");
				lblProgress.setVisible(true);
				btnStop.setVisible(true);
				
				// read dependencies from tree into 2d array list
				for (ExcelFileInfo f : listFileInfo) {
					ArrayList<Dependencies> saveDependencies = new ArrayList<Dependencies>();
					readDependencies(f.dependentCell, saveDependencies);
					relatedFormulaCells.add(saveDependencies);
					
					
				}
				
				for (int i = 0; i < tableResult.getRowCount(); i++) {
					fileName = tableResult.getValueAt(i, 0).toString();
					fileLocation = tableResult.getValueAt(i, 1).toString();
					temp = tableResult.getValueAt(i, 2).toString();
					sheet = temp.substring(0, temp.indexOf("!"));
					cell = temp.substring(temp.indexOf("!"), temp.length());
					
					try {
						File excelFile = new File(fileLocation, fileName);
						replaceString(excelFile, find, replace, sheet, cell, caseSensitive);
						
						if (requestedFormulaAction == ACTION_UPDATE_RELATIVE_FORMULA)
							recalculateFormulas(excelFile, i);
						
						if (terminateThread) {
							//System.out.println("Thread interrupting...");
							lblProgress.setVisible(false);
							btnStop.setVisible(false);
							enableAllButtons();
							terminateThread = false;
							return;
						}
						
						Object replaced = "Yes";
						tableResultModel.setValueAt(replaced, i, 3);
					} catch (InvalidFormatException | IOException e1) {
						// TODO Auto-generated catch block
						//e1.printStackTrace();
						Object replaced = "No";
						tableResultModel.setValueAt(replaced, i, 3);
					}
					
				}
				
			
				break;
				
		}
		
		
		lblProgress.setVisible(false);
		btnStop.setVisible(false);
		enableAllButtons();
		
	}
	
	
	public void setAction(int action, int formulaAction, String location, String find, String replace, boolean includeSubdirectories, boolean caseSensitive, boolean exactCellMatch) {
		if ((action == ACTION_FIND || action == ACTION_REPLACE || action == ACTION_REPLACE_ALL) &&
				 (formulaAction == ACTION_UPDATE_RELATIVE_FORMULA || formulaAction == ACTION_UPDATE_ALL_FORMULA)) {
			requestedAction = action;
			requestedFormulaAction = formulaAction;
			this.location = location;
			this.find = find;
			this.replace = replace;
			this.includeSubdirectories = includeSubdirectories;
			this.caseSensitive = caseSensitive;
			this.exactCellMatch = exactCellMatch;
			
		}
			
		
	}
	
	private void disableAllButtons() {
		buttonFind.setEnabled(false);
		buttonReplace.setEnabled(false);
		buttonReplaceAll.setEnabled(false);
		buttonBrowse.setEnabled(false);
		
		chckbxIncludeSubdirectories.setEnabled(false);
		chckbxCaseSensitive.setEnabled(false);
		chckbxExactCellMatch.setEnabled(false);
		
		rdbtnUpdateRelativeFormula.setEnabled(false);
		rdbtnUpdateAllFormula.setEnabled(false);
		
	}
	
	private void enableAllButtons() {
		buttonFind.setEnabled(true);
		
		int t = tableResult.getSelectedRow();
		if (t != -1)
			buttonReplace.setEnabled(true);
		else 
			buttonReplace.setEnabled(false);
		
		if (hasResult)
			buttonReplaceAll.setEnabled(true);
		buttonBrowse.setEnabled(true);
		
		chckbxIncludeSubdirectories.setEnabled(true);
		chckbxCaseSensitive.setEnabled(true);
		chckbxExactCellMatch.setEnabled(true);
		
		rdbtnUpdateRelativeFormula.setEnabled(true);
		rdbtnUpdateAllFormula.setEnabled(true);
		
	}
	
	private void listSearchResult(File directory, String searchString, boolean includeSubdirectories, boolean caseSensitive, boolean exactCellMatch) {
		ExcelFileInfo[] fileInfo;
		File[] fileList = directory.listFiles();
		
		if (fileList != null)
			for (File fileEntry : fileList) {
				if (fileEntry.isDirectory()) {
					
					if (terminateThread) {
						//System.out.println("Thread interrupting...");
						return;
					}
					
					if (includeSubdirectories) {
						listSearchResult(fileEntry, searchString, includeSubdirectories, caseSensitive, exactCellMatch);
						//System.out.println("Search subdirectories...");
					}
				}
				else {
					try {
						
						if (terminateThread) {
							//System.out.println("Thread interrupting...");
							return;
						}
						
						fileInfo = searchString(searchString, fileEntry, caseSensitive, exactCellMatch);
						for (ExcelFileInfo f : fileInfo) {
							tableResultModel.addRow(new Object[] {f.fileName, f.location, f.sheet + "!" + f.cell, "No"});
							listFileInfo.add(f);
							//System.out.println("list file info add!");
							hasResult = true;
						}
						
					} 
					catch (NotOLE2FileException e) {
						// no need to show error message
						
					}
					catch (InvalidFormatException e) {
						JOptionPane.showMessageDialog(null, "Invalid File Format\n" + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
						
					}
					catch (IOException e) {
						// TODO Auto-generated catch block
						JOptionPane.showMessageDialog(null, "IO Error\n" + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
						//e.printStackTrace();
					}
					catch (RecordFormatException e) {
						JOptionPane.showMessageDialog(null, "Invalid Record Format\n" + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
						
					}
					
					
				
				}
			
			}
		
		
		
	}
		
	private void clearSearchResult() {
		int rows = tableResultModel.getRowCount();
		
		for (int i = rows - 1; i >= 0; i--)
			tableResultModel.removeRow(i);
		
		tableResult.clearSelection();
		
		hasResult = false;
		
		relatedFormulaCells.clear();
		listFileInfo.clear();
	}
			
	private boolean hasHistory(String sheet, String cell) {
		for (Dependencies d: dependenciesHistory) {
			//System.out.println("checking history: sheet=" + d.sheet + ", cell=" + d.cell + " with " + sheet + "!" + cell);
			if (d.sheet.equals(sheet) && d.cell.equals(cell)) {
				//System.out.println("has history!");
				return true;
			}
			
		}
		
		return false;
		
		
	}
	
	private void searchDependencies(String sheetName, String cellReference, File excelFile, TreeNode<Dependencies> root) throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (!FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsx") &&
                !FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsm") &&
                !FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xls")
            )
			return;
				
		Workbook workbook = WorkbookFactory.create(excelFile);
		Sheet worksheet;
			
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			worksheet = workbook.getSheetAt(i);
			
			for (Row row : worksheet) {
				for (Cell cell : row) {
					
					if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						if (cell.getCellFormula().toLowerCase().contains(cellReference.toLowerCase())) {
							Pattern p = Pattern.compile("\\w+");
							Matcher m = p.matcher(cell.getCellFormula().toLowerCase());
							
							while (m.find()) {
							
								if (m.group().equals(cellReference.toLowerCase()))
									if (!hasHistory(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1))) {
										//System.out.println("Found dependencies for " + excelFile.getName() + " > " + sheetName + "!" + cellReference + ": at sheet: " + worksheet.getSheetName() + " at cell: " + CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1));
										root.add(new Dependencies(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
										dependenciesHistory.add(new Dependencies(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
									}
								
							}
						}
						else {
							Pattern p = Pattern.compile("\\w+:\\w+");
							Matcher m = p.matcher(cell.getCellFormula().toLowerCase());
							
							while (m.find()) {
								
								if (CellRangeUtil.contains(CellRangeAddress.valueOf(m.group()), CellRangeAddress.valueOf(cellReference.toLowerCase()))) {
									if (!hasHistory(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1))) {
										//System.out.println("Found celll range dependencies for " + excelFile.getName() + " > " + sheetName + "!" + cellReference + ": at sheet: " + worksheet.getSheetName() + " at cell: " + CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1));
										//System.out.println("cell range: " + m.group());
										root.add(new Dependencies(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
										dependenciesHistory.add(new Dependencies(worksheet.getSheetName(), CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
									}
								}
								
								
							}
						}
							
					}
					
					if (terminateThread) {
						//System.out.println("Thread interrupting...");
						workbook.close();
						return;
					}
					
				}
			}
		}
		
		workbook.close();
		
		for (TreeNode<Dependencies> child: root.getChildren())
			searchDependencies(child.getData().sheet, child.getData().cell, excelFile, child);
		
	}
	
	private void recalculateFormulas(File excelFile, int fileInfoIndex) throws InvalidFormatException, IOException {
		
		//System.out.println("recalculate formulas...");
		
		for (Dependencies d: relatedFormulaCells.get(fileInfoIndex)) {
			
						
			FileInputStream input = new FileInputStream(excelFile);
					
			Workbook workbook = WorkbookFactory.create(input);
			Sheet worksheet = workbook.getSheet(d.sheet);
						
			CellReference ref = new CellReference(d.cell);
						
			Row row = worksheet.getRow(ref.getRow());
			Cell cell = row.getCell(ref.getCol());
						
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
						
			//System.out.println("Start evaluating: " + d.cell);
						
			evaluator.evaluateFormulaCell(cell);
						
			FileOutputStream output = new FileOutputStream(excelFile);
						
			workbook.write(output);
			output.flush();
			output.close();
			workbook.close();
						
		}
					
					
				
	}
	
	
	//Tree Traversal
	private void readDependencies(TreeNode<Dependencies> tree, List<Dependencies> saveList) {
		
		saveList.add(tree.getData());
		for (TreeNode<Dependencies> child : tree.getChildren())
			readDependencies(child, saveList);
		
	}
	
	private ExcelFileInfo[] searchString(String str, File excelFile, boolean caseSensitive, boolean exactCellMatch) throws IOException, InvalidFormatException, NotOLE2FileException, RecordFormatException {
				
		fileInfo.clear();
		
		if (!FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsx") &&
                            !FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsm") &&
                            !FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xls")
                        )
			
			return fileInfo.toArray(new ExcelFileInfo[fileInfo.size()]);
		
				
		Workbook workbook = WorkbookFactory.create(excelFile);
		Sheet worksheet;
			
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			worksheet = workbook.getSheetAt(i);
			
			for (Row row : worksheet) {
				for (Cell cell : row) {
					
					if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
						if (exactCellMatch && caseSensitive) {
							if (cell.getRichStringCellValue().getString().equals(str))
								fileInfo.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
										CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
						}
						else if (exactCellMatch && !caseSensitive) {
							if (cell.getRichStringCellValue().getString().toLowerCase().equals(str.toLowerCase()))
								fileInfo.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
										CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
						}
						else if (!exactCellMatch && caseSensitive) {
							if (cell.getRichStringCellValue().getString().contains(str))
								fileInfo.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
										CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
						}
						else if (!exactCellMatch && !caseSensitive) {
							if (cell.getRichStringCellValue().getString().toLowerCase().contains(str.toLowerCase()))
								fileInfo.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
										CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
						}
					}
				/*	else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
						//System.out.println("Formula result type: " + cell.getCachedFormulaResultType());
						
						if (cell.getCachedFormulaResultType() == Cell.CELL_TYPE_STRING) {
							
							//System.out.println("String value: " + cell.getRichStringCellValue().getString());
							
							if (exactCellMatch && caseSensitive) {
								if (cell.getRichStringCellValue().getString().trim().equals(str.trim()))
//									relatedFormulaCells.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
				//						CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
							}
							else if (exactCellMatch && !caseSensitive) {
								if (cell.getRichStringCellValue().getString().trim().toLowerCase().equals(str.trim().toLowerCase()))
//									relatedFormulaCells.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
			//							CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
							}
							else if (!exactCellMatch && caseSensitive) {
								if (cell.getRichStringCellValue().getString().trim().contains(str.trim()))
//									relatedFormulaCells.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
		//								CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
							}
							else if (!exactCellMatch && !caseSensitive) {
								if (cell.getRichStringCellValue().getString().trim().toLowerCase().contains(str.trim().toLowerCase()))
//									relatedFormulaCells.add(new ExcelFileInfo(excelFile.getName(), excelFile.getParent(), worksheet.getSheetName(), 
	//									CellReference.convertNumToColString(cell.getColumnIndex()) + Integer.toString(cell.getRowIndex() + 1)));
							}
						}
					}*/
					
					
					
				}
				
				
			}
			
		}	
			
		workbook.close();
				
		//System.out.println("Finally the related Formula Cells size: " + relatedFormulaCells.size());
		
		//for (ExcelFileInfo f : relatedFormulaCells)
		//	System.out.println("File: " + f.fileName + "Sheet: " + f.sheet + "  Cell: " + f.cell);
		
		
		return fileInfo.toArray(new ExcelFileInfo[fileInfo.size()]);
	}
	
	private void replaceString(File excelFile, String findString, String replaceString, String sheet, String cellRef, boolean caseSensitive) throws InvalidFormatException, IOException {

		
		if (FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsx") || 
                        FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xlsm") ) {
			
			FileInputStream input = new FileInputStream(excelFile);
			
			XSSFWorkbook workbook = new XSSFWorkbook(input);
			XSSFSheet worksheet = workbook.getSheet(sheet);
			CellReference ref = new CellReference(cellRef);
			XSSFRow row = worksheet.getRow(ref.getRow());
			XSSFCell cell = row.getCell(ref.getCol());
			
			XSSFRichTextString orginalString = cell.getRichStringCellValue();
			String newPlainString;
			
			if (caseSensitive)
				newPlainString = orginalString.getString().replace(findString, replaceString);
			else
				newPlainString = orginalString.getString().replaceAll("(?i)" + findString, replaceString);
			
			
			//System.out.println("orginalString: " + orginalString.getString() + " length: " + orginalString.length());
			//System.out.println("newPlainString: " + newPlainString + " length: " + newPlainString.length());
			
			XSSFRichTextString newString = new XSSFRichTextString(newPlainString);
			
			//System.out.println("newString: " + newString.getString() + " length: " + newString.length());
			
			XSSFFont fontApplied;
			
			for (int i = 0; i < newString.length(); i++) {
				if (i < orginalString.length()) {
					try {
						fontApplied = orginalString.getFontAtIndex(i);
					}
					catch (NullPointerException e) {
						fontApplied = null;
						//System.out.println("null font at: " + i);
					}
						
					if (fontApplied != null)
						newString.applyFont(i, i + 1, fontApplied);
					
					
				}
			}
			
			if (newPlainString.equals(""))
				newString = null;
			
			cell.setCellValue(newString);
			
			
			if (requestedFormulaAction == ACTION_UPDATE_ALL_FORMULA)
				XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
			
						
			FileOutputStream output = new FileOutputStream(excelFile);
			
			workbook.write(output);
			output.flush();
			output.close();
			workbook.close();
			
			
			
		}
		else if (FilenameUtils.getExtension(excelFile.getAbsolutePath()).equalsIgnoreCase("xls")) {
			FileInputStream input = new FileInputStream(excelFile);
			
			HSSFWorkbook workbook = new HSSFWorkbook(input);
			HSSFSheet worksheet = workbook.getSheet(sheet);
			CellReference ref = new CellReference(cellRef);
			HSSFRow row = worksheet.getRow(ref.getRow());
			HSSFCell cell = row.getCell(ref.getCol());
			
			HSSFRichTextString orginalString = cell.getRichStringCellValue();
			String newPlainString;
			
			if (caseSensitive)
				newPlainString = orginalString.getString().replace(findString, replaceString);
			else
				newPlainString = orginalString.getString().replaceAll("(?i)" + findString, replaceString);
			
			
			//System.out.println("orginalString: " + orginalString.getString() + " length: " + orginalString.length());
			//System.out.println("newPlainString: " + newPlainString + " length: " + newPlainString.length());
			
			HSSFRichTextString newString = new HSSFRichTextString(newPlainString);
			
			//System.out.println("newString: " + newString.getString() + " length: " + newString.length());
			
			Short fontApplied;
			
			for (int i = 0; i < newString.length(); i++) {
				if (i < orginalString.length()) {
					try {
						fontApplied = orginalString.getFontAtIndex(i);
					}
					catch (NullPointerException e) {
						fontApplied = null;
						//System.out.println("null font at: " + i);
					}
						
					if (fontApplied != null)
						newString.applyFont(i, i + 1, fontApplied);
					
					
				}
			}
			
			if (newPlainString.equals(""))
				newString = null;
			
			cell.setCellValue(newString);
			
			if (requestedFormulaAction == ACTION_UPDATE_ALL_FORMULA)
				HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
				
						
			FileOutputStream output = new FileOutputStream(excelFile);
			
			workbook.write(output);
			output.flush();
			output.close();
			workbook.close();
			
			
				
			
		}
		else {
			throw new InvalidFormatException("Invalid File Extension.");
		}
		
		
	
		
	}

}
