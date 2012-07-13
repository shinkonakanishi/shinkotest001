package jp.co.snknet.common.excel.controller;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Excel�R���g���[��<br>
 * <br>
 * Excel�̐���
 * 
 * @author Shinko
 * @version 1.0
 */
public class ExcelController {

	private Workbook workBook;		// ���[�N�p�u�b�N

	String msFilePath = "";	// �Ώ�Excel�t�@�C���p�X
	private static final String DEF_SHEET_NAME_DEFUALT = "�V�[�g�P";	// �f�t�H���g�V�[�g���i�V�K�쐬���̂ݎg�p�j
	
    /**
     * �R���X�g���N�^
     */
	public ExcelController() {
		//�@������
		this.initialize();
	}
    /**
     * �R���X�g���N�^
     * 
     * @param filePath �t�@�C���p�X
     */
	public ExcelController(String filePath) {
		//�@������
		this.initialize();
		//�@�l��ێ�
		msFilePath = filePath;
	}
    /**
     * ������
     */
	private void initialize() {
		//
		this.clear();
		//
		msFilePath = "";
	}
    /**
     * �N���A
     */	
	public void clear() {
		// �u�b�N��V�K
		workBook =  new HSSFWorkbook();
		// �V�[�g���쐬
		workBook.createSheet(DEF_SHEET_NAME_DEFUALT);
	}
	/**
	 * �t�@�C���p�X���擾
	 * @return �t�@�C���p�X
	 */
	public String getFilePath() {
		return msFilePath;
	}
    /**
     * ����������
     */
	public void open() throws Exception, IOException {
		
		//
		//�@�N���A
		//
		this.clear();
		
		//
		// Excel���J��
		//
		if (msFilePath.equals("")) {
			//�@�V�K�쐬
			workBook = new HSSFWorkbook();
			
		} else {
			//�@�ҏW
			
			// Excel�t�@�C���̓ǂݍ���
			FileInputStream fis = new FileInputStream(msFilePath);
			POIFSFileSystem fs = new POIFSFileSystem(fis);

			// ���[�N�u�b�N�E�I�u�W�F�N�g�̎擾
			workBook = new HSSFWorkbook(fs);
		}
	}
    /**
     * �t�@�C���֏����o��
     */
	public void save() throws Exception, FileNotFoundException{
		try{
			
			// ���[�N�u�b�N�E�I�u�W�F�N�g���t�@�C���Ƃ��ďo�� 			 		 
			FileOutputStream fileOut = new FileOutputStream(msFilePath);
			workBook.write(fileOut);
			
			//�t�@�C�������
			fileOut.close();

		} catch (Exception ex){
			throw ex;
		}
	}
    /**
     * �t�@�C���֏����o��
     * 
     * @param outputFilePath String �o�̓t�@�C���p�X
     */
	public void save(String outputFilePath) throws Exception, FileNotFoundException{
		// �t�@�C���p�X���㏑���ێ�
		msFilePath = outputFilePath;
		// �t�@�C���֏����o��
		this.save();		
	}
    /**
     * �f�[�^�����
     */
	public void close(){
		// �u�b�N���J��
		workBook = null;
	}
	/**
	 * �V�[�g����ύX
	 * 
	 * @param sheetName String �V�[�g��
	 */
	public void setSheetName(String sheetName) {
		workBook.setSheetName(workBook.getActiveSheetIndex(), sheetName);
	}
	/**
	 * �V�[�g����ύX
	 * 
	 * @param sheetIndex int �V�[�g�C���f�b�N�X
	 * @param sheetName String �V�[�g��
	 */
	public void setSheetName(int sheetIndex, String sheetName) {
		workBook.setSheetName(sheetIndex, sheetName);		
	}
	/**
	 * �V�[�g���폜
	 * 
	 * @param sheetIndex
	 */
	public void deleteSheet(int sheetIndex) {
		workBook.removeSheetAt(sheetIndex);
	}
	/**
	 * �V�[�g�̃N���[�����쐬
	 * 
	 * @param sheetName String �N���[�����̃V�[�g��
	 */
	public void cloneSheet(String sheetName) throws Exception{
		int liSheetIdx = workBook.getSheetIndex(sheetName);
		this.cloneSheet(liSheetIdx);		
	}
	/**
	 * �V�[�g�̃N���[�����쐬
	 * 
	 * @param sheetIndex int �N���[�����̃V�[�g�C���f�b�N�X
	 */
	public void cloneSheet(int sheetIndex) throws Exception{
		// �V�[�g�̃N���[�����쐬
		workBook.cloneSheet(sheetIndex);
		// ����͈͂��R�s�[
		CellRangeAddress lclsRange = ExcelUtility.getPrintArea(ExcelUtility.getPringAreaStringFull(workBook, sheetIndex));
		String lsPrintArea = ExcelUtility.getPrintAreaString(lclsRange);//"$A$1:$F$14";//
		
		workBook.setPrintArea(this.getSheetCount() - 1, lsPrintArea);
	}
	/**
	 * Sheet�R���g���[�����擾
	 * 
	 * @return SheetController �V�[�g�R���g���[��
	 */
	public SheetController getSheet() {
		return this.getSheet(workBook.getActiveSheetIndex());		
	}
	/**
	 * Sheet�R���g���[�����擾
	 * 
	 * @param sheetIndex int �V�[�g�C���f�b�N�X
	 * @return SheetController �V�[�g�R���g���[��
	 */
	public SheetController getSheet(int sheetIndex) {
		return new SheetController(workBook.getSheetAt(sheetIndex));		
	}
	/**
	 * �V�[�g�����擾
	 */
	public int getSheetCount() {
		return workBook.getNumberOfSheets();
	}
	/**
	 * �I���V�[�g�C���f�b�N�X���擾
	 * 
	 * @return int �I�����Ă���V�[�g�̃C���f�b�N�X
	 */
	public int getSelectedSheetIndex() {
		return workBook.getActiveSheetIndex();
	}
	/**
	 * �I���V�[�g�C���f�b�N�X����
	 * 
	 * @param sheetIndex String �I������V�[�g�̃C���f�b�N�X
	 */
	public void setSelectedSheetIndex(int sheetIndex) {
		workBook.setActiveSheet(sheetIndex);
	}
	/**
	 * �Z���ɒl����
	 * 
	 * @param rowIndex
	 * @param column
	 * @param value
	 */
	public void setCellValue(int rowIndex, int columnIndex, String value) {

		SheetController lclsSheet = this.getSheet();
		// �Z���ɒl����
		lclsSheet.setCellValue(rowIndex, columnIndex, value);
	}
	/**
	 * �Z���̒l��String�^�Ŏ擾
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @return
	 */
	public String getCellStringValue(int rowIndex, int columnIndex){
		return this.getSheet().getCellValue(rowIndex, columnIndex);
	}

	
	public void setPrintArea(int aiAddColumnNum) {

		String lsSheetName = "";
		String lsRowStartName = "";
		String lsColumnStartName = "";
		String lsRowEndName = "";
		String lsColumnEndName = "";
		
		String lsPrintArea = workBook.getPrintArea(0);

		int liReadPlace = 0;
		for (int i = 0 ; i < lsPrintArea.length() ; i++) {
			String lsData = lsPrintArea.substring(i, i + 1);
			
			if (lsData.equals("!")
					|| lsData.equals("$")
					|| lsData.equals(":")) {
				liReadPlace ++;
			} else {
				// '�V�[�g��'!$A$1:$C$4
		        switch (liReadPlace) {

		        	case 0:
		        		// �V�[�g��
		        		lsSheetName += lsData;
		        		break;
		        	case 1:
		        		// !$
		        		break;
		        	case 2:
		        		// �J�n�Z����
		        		lsColumnStartName += lsData;		
		        		break;
		        	case 3:
		        		// �J�n�Z���s
		        		lsRowStartName += lsData;		        		
		        		break;
		        	case 4:
		        		// $:
		        		break;
		        	case 5:
		        		// �I���Z����
		        		lsColumnEndName += lsData;
		        		break;
		        	case 6:
		        		// �I���Z���s
		        		lsRowEndName += lsData;
		        		break;
		        	default:
		        }
			}
		}
		
		int liAddRowIndex = Integer.valueOf(lsRowEndName) + aiAddColumnNum;
		
		String lsNewPrintArea = "$" + lsColumnStartName + "$" + lsRowStartName + ":$" + lsColumnEndName + "$" + String.valueOf(liAddRowIndex);
		
		workBook.setPrintArea(0, lsNewPrintArea);
	}
	/**
	 * �񐔂̎擾
	 * 
	 * @return int ��
	 */
	public int getColumnCount() {
		return this.getSheet().getColumnCount();
	}
	/**
	 * �񐔂̎擾
	 * 
	 * @param �V�[�g�C���f�b�N�X
	 * @return int ��
	 */
	public int getColumnCount(int sheetIndex) {
		return this.getSheet(sheetIndex).getColumnCount();
	}
	/**
	 * �s���̎擾
	 * 
	 * @return�@int �s��
	 */
	public int getRowCount() {
		return this.getSheet().getRowCount();
	}
	/**
	 * �s���̎擾
	 * 
	 * @param �V�[�g�C���f�b�N�X
	 * @return�@int �s��
	 */
	public int getRowCount(int sheetIndex) {
		return this.getSheet(sheetIndex).getRowCount();
	}
	public int getPrintRowCount(int sheetIndex) throws Exception{
		CellRangeAddress lclsRange = ExcelUtility.getPrintArea(workBook.getPrintArea(sheetIndex));
		
		return lclsRange.getLastRow() + 1;
	}

	/**
	 * �e���v���[�g�ƂȂ�V�[�g�̓��e���R�s�[
	 * 
	 * @param templateSheetIndex
	 * @param pageNum
	 */
	public void copyTemplateSheet(int templateSheetIndex, int pageNum) throws Exception{
		this.copyTemplateSheet(this.getSelectedSheetIndex(), templateSheetIndex, pageNum);
	}
	/**
	 * �e���v���[�g�ƂȂ�V�[�g�̓��e���R�s�[
	 * 
	 * @param targetSheetIndex
	 * @param templateSheetIndex
	 * @param pageNum
	 */
	public void copyTemplateSheet(int targetSheetIndex, int templateSheetIndex, int pageNum) throws Exception{
		Sheet lclsTargetSheet = workBook.getSheetAt(targetSheetIndex);
		Sheet lclsTemplateSheet = workBook.getSheetAt(templateSheetIndex);

		//
		// �Ώۂ̊J�n�s�C���f�b�N�X���擾
		//
		CellRangeAddress lclsPrintRange = ExcelUtility.getPrintArea(ExcelUtility.getPringAreaStringFull(workBook, templateSheetIndex));
		int liTemplateRowCount = lclsPrintRange.getLastRow() + 1;
		int liTemplateColumnCount = lclsPrintRange.getLastColumn() + 1;
		int liStartRowIndex = (liTemplateRowCount * (pageNum - 1));
		
		//
		//�@�����Z���̃R�s�[
		//
		int liMargedCount = lclsTemplateSheet.getNumMergedRegions();
		
		for (int i = 0 ; i < liMargedCount ; i++) {
			// �����͈͂��擾
			CellRangeAddress lclsRange = lclsTemplateSheet.getMergedRegion(i);
			// �����͈͂���R�s�[��̌����͈̓C���f�b�N�X���擾
			int liFirstRowIndex = lclsRange.getFirstRow() + liStartRowIndex;
			int liFirstColumnIndex = lclsRange.getFirstColumn();
			int liLastRowIndex = lclsRange.getLastRow() + liStartRowIndex;
			int liLastColumnIndex = lclsRange.getLastColumn();
			// �����͈͂�ǉ�
			lclsTargetSheet.addMergedRegion(new CellRangeAddress(liFirstRowIndex, liLastRowIndex, liFirstColumnIndex, liLastColumnIndex));
		}
		
		//
		// �s�̃R�s�[
		//
		
		//
		// �Z���̃R�s�[
		//
		for (int liRow = 0 ; liRow < liTemplateRowCount ; liRow++) {
		
			Row lclsTargetRow = lclsTargetSheet.createRow(liStartRowIndex + liRow);
			Row lclsTempRow = lclsTemplateSheet.getRow(liRow);
					
			if (lclsTempRow != null) {
				for (int liCol = 0 ; liCol < liTemplateColumnCount ; liCol++) {
					Cell lclsTargetCell = lclsTargetRow.createCell(liCol);
					Cell lclsTempCell = lclsTempRow.getCell(liCol);
					
					// �l���擾
					if (lclsTempCell != null) {
						switch (lclsTempCell.getCellType()) {
							case HSSFCell.CELL_TYPE_BLANK :
								break;
							case HSSFCell.CELL_TYPE_BOOLEAN :
								lclsTargetCell.setCellValue(lclsTempCell.getBooleanCellValue());
								break;
							case HSSFCell.CELL_TYPE_ERROR :
								lclsTargetCell.setCellValue(lclsTempCell.getErrorCellValue());
								break;
							case HSSFCell.CELL_TYPE_FORMULA :
								lclsTargetCell.setCellValue(lclsTempCell.getStringCellValue());
								break;
							case HSSFCell.CELL_TYPE_NUMERIC :
								lclsTargetCell.setCellValue(lclsTempCell.getNumericCellValue());
								break;
							case HSSFCell.CELL_TYPE_STRING :
								lclsTargetCell.setCellValue(lclsTempCell.getStringCellValue());
								break;
							default :
						}

						// �X�^�C�����擾
						lclsTargetCell.setCellStyle(lclsTempCell.getCellStyle());

					}
				}			
			}
		}
		
		//
		// ���y�[�W�̑}��
		//
		lclsTargetSheet.setRowBreak(liStartRowIndex + (liTemplateRowCount - 1));

		//
		// ����͈͂̑}��
		//
		this.setPrintArea(liStartRowIndex + liTemplateRowCount);

	}

}