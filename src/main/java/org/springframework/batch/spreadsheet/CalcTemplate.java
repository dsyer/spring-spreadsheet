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

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

/**
 * This utility class provides easy access to processing Open Office Calc worksheets. Code using this
 * only needs to implement the callback interfaces.
 * 
 * @since 11/2/2009
 * @author Greg Turnquist
 * @see ExcelTemplate
 */
public class CalcTemplate {

	private static final Logger logger = Logger.getLogger(CalcTemplate.class);
	
	/**
	 * The file that this instance of CalcTemplate processes.
	 */
	private File file;
	
	/**
	 * Option to skip the first row (usually due to a header being there).
	 */
	private boolean skipFirstRowDefault;

	/**
	 * Standard policy is to NOT skip the first row of a worksheet.
	 */
	public CalcTemplate(File file) {
		this(file, false);
	}

	/**
	 * Set whether or not to skip the first row by default for a particular worksheet.
	 */
	public CalcTemplate(File file, boolean skipFirstRowDefault) {
		this.file = file;
		this.skipFirstRowDefault = skipFirstRowDefault;
	}

	/**
	 * Process each row of the worksheet using the default error handler.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param sheetNum - integer index into the row of the spreadsheet
	 * @param calcCallback - callback defining how to process a row of data
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(int sheetNum, CalcRowCallback<T> calcCallback) {
		return onEachRow(sheetNum, calcCallback, skipFirstRowDefault, new DefaultCalcTemplateErrorHandler<T>());
	}
	
	/**
	 * Process each row of the worksheet, using an alternate error handling strategy.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param sheetNum - integer index into the row of the spreadsheet
	 * @param calcCallback - callback defining how to process a row of data
	 * @param errorHandler - custom error handler
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(int sheetNum, CalcRowCallback<T> calcCallback, CalcTemplateErrorHandler<T> errorHandler) {
		return onEachRow(sheetNum, calcCallback, skipFirstRowDefault, errorHandler);
	}
	
	/**
	 * This is the work horse for row-level worksheet processing.
	 * <p>
	 * 1) Read data from file.<br/>
	 * 2) Find specific worksheet.<br/>
	 * 3) Create an empty List.<br/>
	 * 4) Iterate over the worksheet, building up the list.<br/>
	 * 5) Return the list.
	 * 
	 * @param <T> - type of the object to be returned
	 * @param sheetNum - integer index into the row of the spreadsheet
	 * @param calcCallback - callback defining how to process a row of data
	 * @param skipFirstRow - override default setting of whether or not to skip the first row
	 * @param errorHandler - custom error handler
	 * @return list of T objects
	 */
	public <T> List<T> onEachRow(int sheetNum, CalcRowCallback<T> calcCallback, boolean skipFirstRow, CalcTemplateErrorHandler<T> errorHandler) {
		try {
			SpreadSheet spreadsheet = SpreadSheet.createFromFile(file);
			Sheet sheet = spreadsheet.getSheet(sheetNum);
			
			List<T> results = new ArrayList<T>();
			
			if (skipFirstRow) {
				logger.debug("Skipping first row...");
				for (int row=1; row < sheet.getRowCount(); row++) {
					processRow(calcCallback, results, sheet, row, errorHandler);
				}
			} else {
				logger.debug("Skipping nuthin'!");
				for (int row=0; row < sheet.getRowCount(); row++) {
					processRow(calcCallback, results, sheet, row, errorHandler);
				}
			}

			return results;
			
		} catch (IOException e) {
			e.printStackTrace();
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
	 * @param calcCallback - callback defining how to process a row of data
	 * @param results - list that is being built up by iteration
	 * @param sheet - worksheet that is being processed
	 * @param row - index into spreadsheet row
	 * @param errorHandler - error handler callback
	 */
	private <T> void processRow(CalcRowCallback<T> calcCallback, List<T> results, Sheet sheet, int row, CalcTemplateErrorHandler<T> errorHandler) {
		T rowResult = null;
		try {
			rowResult = calcCallback.mapRow(sheet, row);
		} catch (RuntimeException e) {
			rowResult = errorHandler.handleException(sheet, row, e);
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

}
