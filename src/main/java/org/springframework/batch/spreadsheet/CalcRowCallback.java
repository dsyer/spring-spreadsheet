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

import org.jopendocument.dom.spreadsheet.Sheet;

/**
 * This simple interface defines a callback for processing one row of a spreadsheet.
 * 
 * @since 11/2/2009
 * @author Greg Turnquist
 * @see CalcTemplate
 */
public interface CalcRowCallback<T> {

	/**
	 * This method is used to map a row of data. However, for some use cases, you may need access
	 * to other information about the worksheet, which is why sheet is provided.
	 * 
	 * @param sheet - the entire Calc worksheet
	 * @param row - the specific Calc row being processed
	 * @return a mapped object based on the <code>row</code>, or <code>null</code>
	 */
	public T mapRow(Sheet sheet, int row);
	
}
