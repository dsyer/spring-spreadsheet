/*
 * Copyright 2002-2009 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package org.springframework.batch.spreadsheet;

import java.awt.Point;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

/**
 * This utility class provides easy access to processing Microsoft Office Excel worksheets. Code using this
 * only needs to implement the callback interfaces.
 * 
 * @since 11/2/2009
 * @author Greg Turnquist
 * @see CalcTemplate
 */
public class ExcelTemplate {

	/**
	 * The file that this instance of ExcelTemplate processes.
	 */
	private File file;
	
	/**
	 * Option to skip the first row (usually due to a header being there).
	 */
	private boolean skipFirstRowDefault;

	/**
	 * Standard policy is to NOT skip the first row of a worksheet.
	 */
	public ExcelTemplate(File file) {
		this(file, false);
	}

	/**
	 * Set whether or not to skip the first row by default for a particular worksheet.
	 */
	public ExcelTemplate(File file, boolean skipFirstRowDefault) {
		this.file = file;
		this.skipFirstRowDefault = skipFirstRowDefault;
	}

	/**
	 * Process each row of the worksheet using the default error handler.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param worksheetName - name of the worksheet to process
	 * @param excelCallback - callback defining how to process a row of data
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(String worksheetName, ExcelRowCallback<T> excelCallback) {
		return onEachRow(worksheetName, excelCallback, skipFirstRowDefault, new DefaultExcelTemplateErrorHandler<T>());
	}

	/**
	 * Process each row of the worksheet using a customized error handler.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param worksheetName - name of the worksheet to process
	 * @param excelCallback - callback defining how to process a row of data
	 * @param errorHandler
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(String worksheetName, ExcelRowCallback<T> excelCallback, ExcelTemplateErrorHandler<T> errorHandler) {
		return onEachRow(worksheetName, excelCallback, skipFirstRowDefault, errorHandler);
	}

	/**
	 * This is the work horse for row-level worksheet processing.
	 * <p>
	 * 1) Read data from file.<br/>
	 * 3) Create an empty List.<br/>
	 * 4) Iterate over the specific worksheet, building up the list.<br/>
	 * 5) Return the list.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param worksheetName - name of the worksheet to process
	 * @param excelCallback - callback defining how to process a row of data
	 * @param skipFirstRow
	 * @param errorHandler
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(String worksheetName, ExcelRowCallback<T> excelCallback, boolean skipFirstRow, ExcelTemplateErrorHandler<T> errorHandler) {
		try {
			InputStream inp = new FileInputStream(file);
			HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));

			List<T> results = new ArrayList<T>();

			if (skipFirstRow) {
				boolean firstRow = true;
				for (Row row : wb.getSheet(worksheetName)) {
					if (firstRow) {
						firstRow = false;
						continue;
					}

					processRow(excelCallback, errorHandler, results, row);
				}
			} else {
				for (Row row : wb.getSheet(worksheetName)) {
					processRow(excelCallback, errorHandler, results, row);
				}
			}

			return results;
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * This utility method is used to invoke the row-level callback. It also traps any
	 * runtime exceptions, and runs them through the error handler.
	 * <p>
	 * If the callback returns <code>null</code>, the row is NOT added to the list.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param excelCallback - callback defining how to process a row of data
	 * @param errorHandler
	 * @param results
	 * @param row
	 */
	private <T> void processRow(ExcelRowCallback<T> excelCallback,
			ExcelTemplateErrorHandler<T> errorHandler, List<T> results, Row row) {
		T rowResult = null;
		try {
			rowResult = excelCallback.mapRow(row);
		} catch (RuntimeException e) {
			rowResult = errorHandler.handleException(row, e);
		}
		if (rowResult != null) {
			results.add(rowResult);
		}
	}

	public boolean isSkipFirstRowDefault() {
		return skipFirstRowDefault;
	}

	public void setSkipFirstRowDefault(boolean skipFirstRowDefault) {
		this.skipFirstRowDefault = skipFirstRowDefault;
	}
	
	/**
	 * The following block includes a search algorithm, where you can scan a worksheet for certain expression, and harvest
	 * the cell its found in.
	 * TODO: Write some tests for this!
	 */
	private enum Direction {
		GOING_UP,
		GOING_DOWN,
		TRANSITION_DOWN,
		TRANSITION_ACROSS;
	}

	/**
	 * This function uses a search algorithm to find a cell, and return the coordinates of
	 * that cell. It uses a 2-dimensional search algorithm, scanning (0,0), (1,0), (0,1),
	 * (0,2), (1,1), (2,0), (3,0), (2,1), (1,2), (0,3)...
	 * 
	 *  It is limited by a finite number of steps so it won't run forever. There is
	 *  a default, and user can increase the default, but there is still a limit.
	 *  
	 *  TODO: Test this.
	 *  
	 * @param sheet
	 * @param keyPhrase
	 * @return Point(x=column, y=row)
	 */
	private Point searchForCell(HSSFSheet sheet, String keyPhrase) {
		return searchForCell(sheet, keyPhrase, 1000);
	}

	/**
	 * This is the core method, and used when the user needs to override the number of steps
	 * to take in searching for the cell.
	 * 
	 * TODO: Test this.
	 * 
	 * @param sheet
	 * @param keyPhrase
	 * @param maxSteps
	 * @return
	 */
	private Point searchForCell(HSSFSheet sheet, String keyPhrase, int maxSteps) {
		int row = 0;
		int column = 0;
		Direction parseDirection = Direction.TRANSITION_DOWN;
		int steps = 0;
		while (steps < maxSteps) {
			HSSFCell cell = sheet.getRow(row).getCell(column);
			if (cell != null) {
				if (cell.toString().equals(keyPhrase)) {
					return new Point(column, row);
				}
			}
			switch (parseDirection) {
				case GOING_DOWN:
					row++; column--;
					steps++;
					if (column == 0) parseDirection = Direction.TRANSITION_DOWN;
					break;
				case GOING_UP:
					row--; column++;
					steps++;
					if (row == 0) parseDirection = Direction.TRANSITION_ACROSS;
					break;
				case TRANSITION_DOWN:
					row++;
					steps++;
					parseDirection = Direction.GOING_UP;
					break;
				case TRANSITION_ACROSS:
					column++;
					steps++;
					parseDirection = Direction.GOING_DOWN;
					break;
			}
		}
		throw new RuntimeException("Could not find '" + keyPhrase + "' in less than " + maxSteps + " steps.");
	}


}
