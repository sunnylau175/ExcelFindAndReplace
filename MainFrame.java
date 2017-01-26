import java.awt.Desktop;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.JTextField;
import javax.swing.JLabel;

import java.awt.Font;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JCheckBox;
import javax.swing.JButton;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.UIManager;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.eclipse.wb.swing.FocusTraversalOnArray;

import java.awt.Component;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

import javax.swing.JFileChooser;

import java.io.File;
import java.io.IOException;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

import javax.swing.JRadioButton;
import javax.swing.SwingConstants;

public class MainFrame extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private JTextField textFieldFind;
	private JTextField textFieldReplace;
	private JTable tableResult;
	private JTextField textFieldLocation;
	private JFileChooser folderDialog;
	private DefaultTableModel tableResultModel;
	private JCheckBox chckbxIncludeSubdirectories;
	private JCheckBox chckbxCaseSensitive;
	private JCheckBox chckbxExactCellMatch;
	private SearchEngine excelSearchEngine;
	private JButton buttonFind;
	private JButton buttonReplace;
	private JButton buttonReplaceAll;
	private JButton buttonBrowse;
	private Thread searchEngineThread;
	private JLabel lblProgress;
	private JRadioButton rdbtnUpdateRelativeFormula;
	private JRadioButton rdbtnUpdateAllFormula;
	private int selectedFormulaAction = SearchEngine.ACTION_UPDATE_RELATIVE_FORMULA;
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					MainFrame frame = new MainFrame();
					frame.setVisible(true);
					frame.setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
				} catch (Exception e) {
					JOptionPane.showMessageDialog(null, "Run-time Error\n" + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
					//e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public MainFrame() {
		addWindowListener(new WindowAdapter() {
			@Override
			public void windowClosing(WindowEvent arg0) {
				if (searchEngineThread != null && searchEngineThread.isAlive()) {
					//System.out.println("Window closing... thread alive");
					
					excelSearchEngine.terminateSearchEngine(true);

					// Start the killing thread to let stopping message show properly
					
					KillApp killingThread = new KillApp(searchEngineThread);
					
					//System.out.println("Starting killing thread...");
					
					killingThread.start();
					
				}
				else {
					//System.out.println("Window closing... thread not alive");
					dispose();
				}
				
			}
		});
		
		
		
		folderDialog = new JFileChooser();
		folderDialog.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		
		
		setResizable(false);
		setTitle("ExcelFindAndReplace");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 697, 735);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JCheckBox chckbxTrimFindText = new JCheckBox("Trim Find Text");
		chckbxTrimFindText.setSelected(true);
		chckbxTrimFindText.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		chckbxTrimFindText.setBounds(14, 220, 186, 23);
		contentPane.add(chckbxTrimFindText);
		
		JCheckBox chckbxTrimReplaceText = new JCheckBox("Trim Replace Text");
		chckbxTrimReplaceText.setSelected(true);
		chckbxTrimReplaceText.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		chckbxTrimReplaceText.setBounds(202, 220, 186, 23);
		contentPane.add(chckbxTrimReplaceText);
				
		textFieldFind = new JTextField();
		textFieldFind.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		textFieldFind.setBounds(14, 97, 517, 27);
		contentPane.add(textFieldFind);
		textFieldFind.setColumns(10);
		
		textFieldReplace = new JTextField();
		textFieldReplace.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if ((searchEngineThread == null || !searchEngineThread.isAlive()) && tableResultModel.getRowCount() > 0) {
					
					//System.out.println("Creating thread...");
										
					excelSearchEngine.setAction(SearchEngine.ACTION_REPLACE_ALL, selectedFormulaAction, textFieldLocation.getText(), chckbxTrimFindText.isSelected() ? textFieldFind.getText().trim() : textFieldFind.getText(), chckbxTrimReplaceText.isSelected() ? textFieldReplace.getText().trim() : textFieldReplace.getText(), chckbxIncludeSubdirectories.isSelected(), chckbxCaseSensitive.isSelected(), chckbxExactCellMatch.isSelected());
					searchEngineThread = new Thread(excelSearchEngine);
					searchEngineThread.start();
				}
				
			}
		});
		textFieldReplace.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		textFieldReplace.setColumns(10);
		textFieldReplace.setBounds(14, 160, 517, 27);
		contentPane.add(textFieldReplace);
		
		JLabel lblFind = new JLabel("Find:");
		lblFind.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		lblFind.setBounds(14, 72, 46, 15);
		contentPane.add(lblFind);
		
		JLabel lblReplace = new JLabel("Replace:");
		lblReplace.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		lblReplace.setBounds(14, 135, 73, 15);
		contentPane.add(lblReplace);
		
		chckbxIncludeSubdirectories = new JCheckBox("Include Subdirectories");
		chckbxIncludeSubdirectories.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		chckbxIncludeSubdirectories.setBounds(14, 193, 186, 23);
		contentPane.add(chckbxIncludeSubdirectories);
		
		buttonFind = new JButton("Find");
		buttonFind.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (searchEngineThread == null || !searchEngineThread.isAlive()) {
					excelSearchEngine.setAction(SearchEngine.ACTION_FIND, selectedFormulaAction, textFieldLocation.getText(), chckbxTrimFindText.isSelected() ? textFieldFind.getText().trim() : textFieldFind.getText(), chckbxTrimReplaceText.isSelected() ? textFieldReplace.getText().trim() : textFieldReplace.getText(), chckbxIncludeSubdirectories.isSelected(), chckbxCaseSensitive.isSelected(), chckbxExactCellMatch.isSelected());
					searchEngineThread = new Thread(excelSearchEngine);
					searchEngineThread.start();
				}
			}
		});
		buttonFind.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		buttonFind.setBounds(559, 96, 116, 27);
		contentPane.add(buttonFind);
		
		buttonFind.getRootPane().setDefaultButton(buttonFind);
		
		JScrollPane scrollPaneResult = new JScrollPane();
		scrollPaneResult.setBounds(14, 314, 661, 376);
		contentPane.add(scrollPaneResult);
		
		tableResult = new JTable();
		tableResult.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				
				if (arg0.getClickCount() == 2) {
					String fileName, fileLocation;
					JTable table = (JTable)arg0.getSource();
					
					fileName = table.getValueAt(tableResult.getSelectedRow(), 0).toString();
					fileLocation = table.getValueAt(tableResult.getSelectedRow(), 1).toString();
					
					File f = new File(fileLocation, fileName);
					
					
					try {
						Desktop.getDesktop().edit(f);
					} catch (IOException e) {
						try {
							Desktop.getDesktop().open(f);
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							//e1.printStackTrace();
							JOptionPane.showMessageDialog(null, "File Editing Error: \n" + f.getAbsolutePath(), "Error Message", JOptionPane.ERROR_MESSAGE);
						}
																		
						//e.printStackTrace();
					}
					
				}
			}
		});
		tableResult.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		tableResult.setBorder(null);
		tableResultModel = new DefaultTableModel(
			new Object[][] {
				{null, null, null, null},
			},
			new String[] {
				"File Name", "Location", "Cell", "Replaced"
			}
		) {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;
			boolean[] columnEditables = new boolean[] {
				false, false, false, false
			};
			public boolean isCellEditable(int row, int column) {
				return columnEditables[column];
			}
		};
		
		tableResult.setModel(tableResultModel);
		
		tableResult.getColumnModel().getColumn(0).setPreferredWidth(192);
		tableResult.getColumnModel().getColumn(1).setPreferredWidth(270);
		tableResult.getColumnModel().getColumn(2).setPreferredWidth(120);
		
		tableResultModel.removeRow(0);
		tableResult.setRowHeight(tableResult.getRowHeight() + 6);
		
		tableResult.getTableHeader().setFont(new Font("PMingLiU", Font.PLAIN, 16));
		scrollPaneResult.setViewportView(tableResult);
		
		tableResult.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
			public void valueChanged(ListSelectionEvent event) {
				
				if (tableResult.getSelectedRow() != -1) {
					if (searchEngineThread == null || !searchEngineThread.isAlive())
						buttonReplace.setEnabled(true);
					
				}
				else 
					buttonReplace.setEnabled(false);
				
			}
						
			
		}); 
		
		tableResult.setShowVerticalLines(false);
		tableResult.setShowHorizontalLines(false);
		
		tableResult.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		
		buttonReplace = new JButton("Replace");
		buttonReplace.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if (searchEngineThread == null || !searchEngineThread.isAlive()) {
					excelSearchEngine.setAction(SearchEngine.ACTION_REPLACE, selectedFormulaAction, textFieldLocation.getText(), chckbxTrimFindText.isSelected() ? textFieldFind.getText().trim() : textFieldFind.getText(), chckbxTrimReplaceText.isSelected() ? textFieldReplace.getText().trim() : textFieldReplace.getText(), chckbxIncludeSubdirectories.isSelected(), chckbxCaseSensitive.isSelected(), chckbxExactCellMatch.isSelected());
					searchEngineThread = new Thread(excelSearchEngine);
					searchEngineThread.start();
				}
			}
		});
		buttonReplace.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		buttonReplace.setBounds(559, 135, 116, 27);
		contentPane.add(buttonReplace);
		
		buttonReplace.setEnabled(false);
		
		buttonReplaceAll = new JButton("Replace All");
		buttonReplaceAll.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (searchEngineThread == null || !searchEngineThread.isAlive()) {
					excelSearchEngine.setAction(SearchEngine.ACTION_REPLACE_ALL, selectedFormulaAction, textFieldLocation.getText(), chckbxTrimFindText.isSelected() ? textFieldFind.getText().trim() : textFieldFind.getText(), chckbxTrimReplaceText.isSelected() ? textFieldReplace.getText().trim() : textFieldReplace.getText(), chckbxIncludeSubdirectories.isSelected(), chckbxCaseSensitive.isSelected(), chckbxExactCellMatch.isSelected());
					searchEngineThread = new Thread(excelSearchEngine);
					searchEngineThread.start();
				}
			}
		});
		buttonReplaceAll.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		buttonReplaceAll.setBounds(559, 174, 116, 27);
		contentPane.add(buttonReplaceAll);
		
		buttonReplaceAll.setEnabled(false);
		
		JLabel lblLocation = new JLabel("Location:");
		lblLocation.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		lblLocation.setBounds(14, 10, 109, 15);
		contentPane.add(lblLocation);
		
		textFieldLocation = new JTextField();
		textFieldLocation.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		textFieldLocation.setColumns(10);
		textFieldLocation.setBounds(14, 35, 517, 27);
		contentPane.add(textFieldLocation);
		
		buttonBrowse = new JButton("Browse");
		buttonBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (folderDialog.showOpenDialog(null) == JFileChooser.APPROVE_OPTION)
					textFieldLocation.setText(folderDialog.getSelectedFile().toString());
				
			}
		});
		buttonBrowse.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		buttonBrowse.setBounds(559, 34, 116, 27);
		contentPane.add(buttonBrowse);
		
		chckbxCaseSensitive = new JCheckBox("Case Sensitive");
		chckbxCaseSensitive.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		chckbxCaseSensitive.setBounds(202, 194, 137, 23);
		contentPane.add(chckbxCaseSensitive);
		
		chckbxExactCellMatch = new JCheckBox("Exact Cell Match");
		chckbxExactCellMatch.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		chckbxExactCellMatch.setBounds(342, 194, 137, 23);
		contentPane.add(chckbxExactCellMatch);
		contentPane.setFocusTraversalPolicy(new FocusTraversalOnArray(new Component[]{textFieldLocation, buttonBrowse, textFieldFind, textFieldReplace, chckbxIncludeSubdirectories, buttonFind, buttonReplace, buttonReplaceAll, lblLocation, lblFind, lblReplace, scrollPaneResult, tableResult}));
		
		
		lblProgress = new JLabel("Searching dependencies...");
		lblProgress.setHorizontalAlignment(SwingConstants.RIGHT);
		lblProgress.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		lblProgress.setBounds(470, 217, 175, 15);
		
				
		contentPane.add(lblProgress);
		lblProgress.setVisible(false);
				
		JButton btnStop = new JButton("");
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				excelSearchEngine.terminateSearchEngine(true);
				
			}
		});
		btnStop.setBounds(648, 211, 28, 28);
		
		ImageIcon stopIcon = new ImageIcon(this.getClass().getResource("stop.png"));
		btnStop.setIcon(stopIcon);
		btnStop.setBorder(BorderFactory.createEmptyBorder());
		btnStop.setContentAreaFilled(false);
		
		contentPane.add(btnStop);

		btnStop.setVisible(false);
		
		
		rdbtnUpdateRelativeFormula = new JRadioButton("Update all dependent formula cells");
		rdbtnUpdateRelativeFormula.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				selectedFormulaAction = excelSearchEngine.ACTION_UPDATE_RELATIVE_FORMULA;
				//System.out.println("selected action: relative");
			}
		});
		rdbtnUpdateRelativeFormula.setSelected(true);
		rdbtnUpdateRelativeFormula.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		rdbtnUpdateRelativeFormula.setBounds(14, 248, 368, 23);
		contentPane.add(rdbtnUpdateRelativeFormula);
		
		rdbtnUpdateAllFormula = new JRadioButton("Update all formula cells (low performance)");
		rdbtnUpdateAllFormula.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				selectedFormulaAction = excelSearchEngine.ACTION_UPDATE_ALL_FORMULA;
				//System.out.println("selected action: all");
			}
		});
		rdbtnUpdateAllFormula.setFont(new Font("PMingLiU", Font.PLAIN, 16));
		rdbtnUpdateAllFormula.setBounds(14, 275, 321, 23);
		
		
		ButtonGroup bg = new ButtonGroup();
		
		bg.add(rdbtnUpdateRelativeFormula);
		bg.add(rdbtnUpdateAllFormula);
		contentPane.add(rdbtnUpdateAllFormula);
		
		excelSearchEngine = new SearchEngine(tableResult, buttonFind, buttonReplace, buttonReplaceAll, buttonBrowse, lblProgress, btnStop, chckbxIncludeSubdirectories, chckbxCaseSensitive, chckbxExactCellMatch, rdbtnUpdateRelativeFormula, rdbtnUpdateAllFormula);
		
		
		
	}
}
