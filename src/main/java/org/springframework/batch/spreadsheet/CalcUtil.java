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

import org.jdom.Attribute;
import org.jopendocument.dom.spreadsheet.MutableCell;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;

/**
 * This class provides some convenience functions to help with processing Open Office Calc spreadsheets.
 * 
 * @since 11/2/2009
 * @author Greg Turnquist
 * @see CalcTemplate
 */
public class CalcUtil {

	/**
	 * Fetches the value of a cell found at (row, column).
	 * 
	 * @param sheet - worksheet where the cell is located
	 * @param column
	 * @param row
	 * @return the value stored at the given cell coordinates
	 */
	public static String getAttr(Sheet sheet, int column, int row) {
		MutableCell<SpreadSheet> cell = sheet.getCellAt(column, row);
		String results = cell.getValue().toString();
		if (results == null || results.equals("")) {
			results = ((Attribute)cell.getElement().getAttributes().get(0)).getValue();
		}
		return results.equals("") ? null : results;
	}

}
