import java.awt.BorderLayout;
import java.awt.Container;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.GridBagLayout;
import java.awt.HeadlessException;

import javax.swing.JFrame;
import javax.swing.JLabel;

import java.awt.Font;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JPanel;
import javax.swing.ScrollPaneConstants;
import javax.swing.border.LineBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.lang.Object;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.awt.Color;

import javax.swing.JButton;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JTable;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import javax.swing.JTree;
import javax.swing.JTextArea;
import java.awt.SystemColor;

public class win_CompareExcel {

	private JFrame ExcelFrame;
	private JTable tbl_excel1;
	private JTable tbl_excel2;
	private JLabel lbl_file1;
	private JLabel lbl_file2;
	public static String selectedFileName1;
	private JScrollPane scrollPane;
	public static String selectedFileName2;
	private JScrollPane scrollPane1;
	private JButton btnCompare;
	public String FileName1 = "";
	public String FileName2 = "";
	public static int DSheetNo = 0;
	public static int DRowNo = 0;
	public static int DColNo = 0;
	public static String lblerror = "";
	public JTextArea lbl_result;
	public static JButton btnFormat;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					win_CompareExcel window = new win_CompareExcel();
					window.ExcelFrame.setVisible(true);

				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public win_CompareExcel() {
		initialize();
		btnFormat.setVisible(false);
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		ExcelFrame = new JFrame();
		ExcelFrame.setBounds(100, 100, 450, 300);
		ExcelFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		ExcelFrame.setExtendedState(ExcelFrame.getExtendedState()
				| JFrame.MAXIMIZED_BOTH);
		ExcelFrame.getContentPane().setLayout(null);

		JLabel lblCompareExcelV = new JLabel("Compare Excel V 1.0");
		lblCompareExcelV.setBounds(24, 11, 229, 52);
		lblCompareExcelV.setFont(new Font("Tahoma", Font.PLAIN, 24));
		ExcelFrame.getContentPane().add(lblCompareExcelV);

		JLabel lblSelectFile = new JLabel("Select File1 : ");
		lblSelectFile.setBounds(34, 74, 118, 14);
		ExcelFrame.getContentPane().add(lblSelectFile);

		JLabel lblNewLabel = new JLabel("Select File2 :");
		lblNewLabel.setBounds(36, 111, 87, 14);
		ExcelFrame.getContentPane().add(lblNewLabel);
		/**
		 * File Browser1
		 */
		JButton btn_excel1 = new JButton("Browse");
		btn_excel1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				btnFormat.setVisible(false);
				try {
					JFileChooser ExcelChooser1 = new JFileChooser();

					ExcelChooser1.setFileFilter(new FileFilter() {

						public String getDescription() {
							return "Excel files";
						}

						public boolean accept(File f) {
							if (f.isDirectory()) {
								return true;
							} else {
								String filename = f.getName().toLowerCase();
								return filename.endsWith(".xls")
										|| filename.endsWith(".xlsx");
							}
						}
					});
					int result = ExcelChooser1.showOpenDialog(ExcelChooser1);
					if (result == JFileChooser.APPROVE_OPTION) {
						File selectedFile = ExcelChooser1.getSelectedFile();

						selectedFileName1 = selectedFile.getAbsolutePath()
								.toString();
						lbl_file1.setText(selectedFile.getAbsolutePath()
								.toString());
						lbl_file1.setToolTipText(selectedFile.getAbsolutePath()
								.toString());

					} else {
						lbl_file1.setText("Please Select a File");
					}
				} catch (HeadlessException e) {
				}
			}
		});
		btn_excel1.setBounds(108, 70, 87, 23);
		ExcelFrame.getContentPane().add(btn_excel1);
		/**
		 * File Browser2
		 */
		JButton btn_excel2 = new JButton("Browse");
		btn_excel2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				btnFormat.setVisible(false);
				try {
					JFileChooser ExcelChooser2 = new JFileChooser();

					ExcelChooser2.setFileFilter(new FileFilter() {

						public String getDescription() {
							return "Excel files";
						}

						public boolean accept(File f) {
							if (f.isDirectory()) {
								return true;
							} else {
								String filename = f.getName().toLowerCase();
								return filename.endsWith(".xls")
										|| filename.endsWith(".xlsx");
							}
						}
					});
					int result = ExcelChooser2.showOpenDialog(ExcelChooser2);
					if (result == JFileChooser.APPROVE_OPTION) {
						File selectedFile = ExcelChooser2.getSelectedFile();

						selectedFileName2 = selectedFile.getAbsolutePath()
								.toString();
						lbl_file2.setText(selectedFile.getAbsolutePath()
								.toString());
						lbl_file2.setToolTipText(selectedFile.getAbsolutePath()
								.toString());

					} else {
						lbl_file1.setText("Please Select a File");
					}
				} catch (HeadlessException e) {
				}
			}
		});
		btn_excel2.setBounds(108, 107, 87, 23);
		ExcelFrame.getContentPane().add(btn_excel2);

		lbl_file1 = new JLabel("");
		lbl_file1.setBounds(226, 74, 476, 14);
		ExcelFrame.getContentPane().add(lbl_file1);

		lbl_file2 = new JLabel("");
		lbl_file2.setBounds(226, 111, 492, 14);
		ExcelFrame.getContentPane().add(lbl_file2);
		/**
		 * File Compare Action button
		 */
		btnCompare = new JButton("Compare");
		btnCompare.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				btnFormat.setVisible(false);
				lblerror = "";
				FileName1 = lbl_file1.getText();
				FileName2 = lbl_file2.getText();
				if (FileName1.isEmpty() || FileName2.isEmpty()) {
					JOptionPane.showMessageDialog(null,
							"Please Select Two Excel Files");
				} else {
					CompareExcel(FileName1, FileName2);
					lbl_result.setText(lblerror);
				}
			}
		});
		btnCompare.setBounds(34, 167, 89, 23);
		ExcelFrame.getContentPane().add(btnCompare);

		lbl_result = new JTextArea();
		lbl_result.setBackground(SystemColor.control);
		lbl_result.setBounds(658, 31, 549, 647);
		ExcelFrame.getContentPane().add(lbl_result);
		/**
		 * File Formatting
		 */
		btnFormat = new JButton("Format");
		btnFormat.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				GetExcelData(FileName1);
				List<String> File1CellValue = new ArrayList<String>();
				List<String> File2CellValue = new ArrayList<String>();
				List<String> File1DataAddress = new ArrayList<String>();
				List<String> File2DataAddress = new ArrayList<String>();
				File1CellValue = GetExcelData(FileName1);
				File2CellValue = GetExcelData(FileName2);
				File1DataAddress = GetDataAddress(FileName1);
				File2DataAddress = GetDataAddress(FileName2);
				for (int i = 0; i <= File1CellValue.size() - 1; i++) {
					if (!File1CellValue.get(i).trim()
							.equals(File2CellValue.get(i).trim()))

					{

						FormatCell(FileName2, File2DataAddress.get(i));

					}
				}
				JOptionPane.showMessageDialog(null, "File2 formatted");
				btnFormat.setVisible(false);
			}
		});
		btnFormat.setBounds(34, 213, 89, 23);
		ExcelFrame.getContentPane().add(btnFormat);

	}
	
	/**
	 * File Compare Method
	 */

	public static void CompareExcel(String filePath1, String filePath2) {
		boolean FileNameComp = false;
		boolean SheetNoComp = false;
		boolean SheetNameComp = false;
		boolean SheetRowComp = false;
		boolean SheetColComp = false;
		boolean SheetDataComp = false;
		boolean SheetAddsComp = false;
		String CellAdss = "";

		if (filePath1.equals(filePath2)) {
			FileNameComp = true;

		} else {
			FileNameComp = false;
		}
		if (FileNameComp) {
			lblerror = lblerror + "\n1. File Names are Equal";
		} else {
			lblerror = lblerror + "\n1. File Names are Not Equal";
		}

		if (GetExcelSheetNos(filePath1) != GetExcelSheetNos(filePath2)) {
			SheetNoComp = false;

		} else {
			SheetNoComp = true;

		}

		if (SheetNoComp) {
			lblerror = lblerror + "\n2. Sheet Numbers in the Files are Equal";
			String[] File1SheetNames = new String[GetExcelSheetNos(filePath1)];
			String[] File2SheetNames = new String[GetExcelSheetNos(filePath2)];
			File1SheetNames = GetExcelSheetNames(filePath1);
			File2SheetNames = GetExcelSheetNames(filePath2);
			SheetNameComp = true;
			for (int i = 0; i <= GetExcelSheetNos(filePath1) - 1; i++) {

				if (!File1SheetNames[i].equals(File2SheetNames[i]))

				{
					SheetNameComp = false;

				}
			}
		} else {
			lblerror = lblerror
					+ "\n2. Sheet Numbers in the Files are Not Equal";
		}

		if (SheetNameComp) {
			lblerror = lblerror + "\n3. Excel Sheet Names are equal";
			List<Integer> File1RowCount = GetPhysicalRows(filePath1);
			List<Integer> File2RowCount = GetPhysicalRows(filePath2);
			SheetRowComp = true;
			for (int i = 0; i <= GetExcelSheetNos(filePath1) - 1; i++) {
				if (File1RowCount.get(i) != File2RowCount.get(i))

				{
					SheetRowComp = false;

				}
			}

		} else {
			lblerror = lblerror + "\n3. Excel Sheet Names are not equal";
		}
		if (SheetRowComp) {
			lblerror = lblerror + "\n4. Number of rows in the sheets are equal";
			List<Integer> File1ColCount = GetPhysicalCols(filePath1);
			List<Integer> File2ColCount = GetPhysicalCols(filePath2);
			SheetColComp = true;
			for (int i = 0; i <= File1ColCount.size() - 1; i++) {
				if (!File1ColCount.get(i).equals(File2ColCount.get(i)))

				{
					SheetColComp = false;

				}
			}

		} else {
			lblerror = lblerror
					+ "\n4. Number of rows in the sheets are not equal";
		}
		if (SheetColComp) {
			lblerror = lblerror
					+ "\n5. Number of Columns in the sheets are equal";
			GetExcelData(filePath1);
			List<String> File1CellValue = new ArrayList<String>();
			List<String> File2CellValue = new ArrayList<String>();
			List<String> File1DataAddress = new ArrayList<String>();
			List<String> File2DataAddress = new ArrayList<String>();
			File1CellValue = GetExcelData(filePath1);
			File2CellValue = GetExcelData(filePath2);
			File1DataAddress = GetDataAddress(filePath1);
			File2DataAddress = GetDataAddress(filePath2);
			SheetDataComp = true;
			SheetAddsComp = true;
			for (int i = 0; i <= File1CellValue.size() - 1; i++) {
				if (!File1CellValue.get(i).trim()
						.equals(File2CellValue.get(i).trim()))

				{

					CellAdss = CellAdss + "\n" + File2DataAddress.get(i);
					SheetDataComp = false;
					SheetAddsComp = false;
					btnFormat.setVisible(true);

				}
			}

		} else {
			lblerror = lblerror
					+ "\n5. Number of Columns in the sheets are not equal";
		}
		if (SheetDataComp) {
			lblerror = lblerror + "\n6. File Data is equal";
		} else {
			lblerror = lblerror
					+ "\n6. File Data is NOT equal in the following Cells Click 'Format' to format the cells "
					+ CellAdss;

		}

	}
	/**
	 * Get Sheet Number Method
	 */
	public static int GetExcelSheetNos(String filePath) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}
		int SheetNos = workbook.getNumberOfSheets();
		return SheetNos;
	}
	/**
	 * Get Sheet Names Method
	 */
	public static String[] GetExcelSheetNames(String filePath) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}
		int SheetNos = workbook.getNumberOfSheets();
		String[] SheetNames = new String[SheetNos];
		for (int i = 0; i <= SheetNos - 1; i++) {
			SheetNames[i] = workbook.getSheetName(i);
		}
		return SheetNames;
	}
	/**
	 * Get Row Numbers Method
	 */
	public static List<Integer> GetPhysicalRows(String filePath) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}
		int SheetCount = workbook.getNumberOfSheets();
		List<Integer> RowCount = new ArrayList<Integer>();
		for (int i = 0; i <= SheetCount - 1; i++) {
			sheet = workbook.getSheetAt(i);
			RowCount.add(sheet.getPhysicalNumberOfRows());

		}

		return RowCount;
	}
	/**
	 * Get Columns Number Method
	 */
	public static List<Integer> GetPhysicalCols(String filePath) {
		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}
		int SheetCount = workbook.getNumberOfSheets();
		List<Integer> RowCount = new ArrayList<Integer>();
		List<Integer> ColCount = new ArrayList<Integer>();
		for (int i = 0; i <= SheetCount - 1; i++) {
			sheet = workbook.getSheetAt(i);
			RowCount.add(sheet.getPhysicalNumberOfRows());
			for (int j = 0; j <= RowCount.size() - 1; j++) {
				Row SRow = sheet.getRow(j);
				try {
					ColCount.add(SRow.getPhysicalNumberOfCells());
				} catch (Exception e) {
				}
			}

		}
		return ColCount;

	}
	/**
	 * Get Excel Data Method
	 */
	public static List<String> GetExcelData(String FileName) {
		String filePath = FileName;

		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}

		int SheetNos = workbook.getNumberOfSheets();
		DSheetNo = SheetNos;
		List<String> CellValue = new ArrayList<String>();
		List<String> DataAddress = new ArrayList<String>();
		for (int SL = 0; SL <= SheetNos - 1; SL++) {
			sheet = workbook.getSheetAt(SL);
			int RCount = sheet.getPhysicalNumberOfRows();
			DRowNo = RCount;
			for (int RL = 0; RL <= RCount - 1; RL++) {
				Row row = sheet.getRow(RL);
				int CCount = row.getPhysicalNumberOfCells();
				DColNo = CCount;
				for (int CL = 0; CL <= CCount - 1; CL++) {
					CellValue.add(row.getCell(CL).toString());

				}
			}
		}

		return CellValue;

	}

	/**
	 * Get Cell Address Method
	 */
	public static List<String> GetDataAddress(String FileName) {
		String filePath = FileName;

		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);
			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}

		int SheetNos = workbook.getNumberOfSheets();
		DSheetNo = SheetNos;
		List<String> CellValue = new ArrayList<String>();
		List<String> DataAddress = new ArrayList<String>();
		for (int SL = 0; SL <= SheetNos - 1; SL++) {
			sheet = workbook.getSheetAt(SL);
			int RCount = sheet.getPhysicalNumberOfRows();
			DRowNo = RCount;
			for (int RL = 0; RL <= RCount - 1; RL++) {
				Row row = sheet.getRow(RL);
				int CCount = row.getPhysicalNumberOfCells();
				DColNo = CCount;
				for (int CL = 0; CL <= CCount - 1; CL++) {
					DataAddress.add(sheet.getSheetName() + ","
							+ row.getRowNum() + "," + CL);

				}
			}
		}

		return DataAddress;

	}
	/**
	 * Format File Method
	 */
	public static void FormatCell(String FileName, String CellAdds) {
		String filePath = FileName;

		FileInputStream fis = null;
		try {
			fis = new FileInputStream(filePath);
		} catch (FileNotFoundException e) {
		}
		Workbook workbook = null;
		Sheet sheet;
		if (filePath.endsWith(".xlsx"))
			try {
				workbook = new XSSFWorkbook(fis);

			} catch (IOException e) {
			}
		else if (filePath.endsWith(".xls"))
			workbook = new HSSFWorkbook();
		else {
			System.err.println("Invalid file type");
		}
		String SheetName = CellAdds.split(",")[0];
		String RowID = CellAdds.split(",")[1];
		String ColID = CellAdds.split(",")[2];
		sheet = workbook.getSheet(SheetName);
		CellStyle style = workbook.createCellStyle();
		org.apache.poi.ss.usermodel.Font font = workbook.createFont();
		font.setFontName(HSSFFont.FONT_ARIAL);
		font.setFontHeightInPoints((short) 11);
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(HSSFColor.BLUE.index);
		style.setFont(font);
		Row row = sheet.getRow(Integer.parseInt(RowID));
		Cell cell = row.getCell(Integer.parseInt(ColID));
		cell.setCellStyle(style);
		sheet.autoSizeColumn((short) 1);
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(new File(FileName));
			workbook.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (fos != null) {
				try {
					fos.flush();
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}

		}
	}
}
