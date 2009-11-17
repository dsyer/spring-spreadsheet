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
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.jopendocument.dom.spreadsheet.Sheet;
import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.springframework.batch.spreadsheet.support.EmptyPhoneBookEntry;
import org.springframework.batch.spreadsheet.support.PhoneBookEntry;


/**
 * @author Greg Turnquist
 */
public class TestCalcTemplate {
	
	private static final Logger logger = Logger.getLogger(TestCalcTemplate.class);
	
	private String pathname = "src" + File.separator + "test" + File.separator + "resources";
		
	// TODO: Workaround (11/2/2009 GLT) - Replace this method with a real Calc spreadsheet.
	@Before
	public void copySpreadsheets() throws FileNotFoundException, IOException {
		logger.debug("TODO: Workaround (11/2/2009 GLT) - Replace this method with a real Calc spreadsheet.");
		logger.debug("Porting Excel spreadsheet to Calc...");
		File input = new File(pathname + File.separator + "phonebook_with_holes.xls");
		
		ExcelTemplate et = new ExcelTemplate(input);
		List<PhoneBookEntry> entries = et.onEachRow("Sheet1", new ExcelRowCallback<PhoneBookEntry>() {
			public PhoneBookEntry mapRow(Row row) {
				String name = "";
				try { name = row.getCell(0).getStringCellValue(); } catch (Exception e) {}
				
				String address = "";
				try { address = row.getCell(1).getStringCellValue(); } catch (Exception e) {}
				
				String phone = "";
				try { phone = row.getCell(2).getStringCellValue(); } catch (Exception e) {}
				
				return new PhoneBookEntry(name, address, phone);
			}
		});

		PhoneBookEntry header = entries.get(0);
		String[] columns = new String[]{header.getName(), header.getAddress(), header.getPhone()};

		final Object[][] data = new Object[entries.size()-1][3];
		
		for (int i=1; i < entries.size(); i++) {
			PhoneBookEntry entry = entries.get(i);
			logger.debug("Adding " + entry + " to the items to be stored in this spreadsheet.");
			data[i-1] = new Object[]{entry.getName(), entry.getAddress(), entry.getPhone()};
		}

		TableModel model = new DefaultTableModel(data, columns);
		
		final File file = new File(pathname + File.separator + "phonebook_with_holes.ods");
		SpreadSheet.createEmpty(model).saveAs(file);
		logger.debug("Done porting file.");
	}
	
	@Test
	public void testReadingSimpleCalcSpreadsheet() throws IOException {
		File file = new File(pathname + File.separator + "phonebook.ods");		
		CalcTemplate ct = new CalcTemplate(file);
		
		List<PhoneBookEntry> results = 
			ct.onEachRow(0, new CalcRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Sheet sheet, int row) {
					return new PhoneBookEntry(
							CalcUtil.getAttr(sheet, 0, row),
							CalcUtil.getAttr(sheet, 1, row),
							CalcUtil.getAttr(sheet, 2, row));
				}
			});
		
		Assert.assertEquals(2, results.size());
		
		Assert.assertEquals("Name", results.get(0).getName());
		Assert.assertEquals("Address", results.get(0).getAddress());
		Assert.assertEquals("Phone", results.get(0).getPhone());
		
		Assert.assertEquals("Peter Gibbons", results.get(1).getName());
		Assert.assertEquals("123 ABC Drive", results.get(1).getAddress());
		Assert.assertEquals("555-821-2123", results.get(1).getPhone());
	}
	
	@Test
	public void testReadingSimpleCalcSpreadsheetSkippingHeader() {
		File file = new File(pathname + File.separator + "phonebook.ods");		
		CalcTemplate et = new CalcTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow(0, new CalcRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Sheet sheet, int row) {
					return new PhoneBookEntry(
							CalcUtil.getAttr(sheet, 0, row),
							CalcUtil.getAttr(sheet, 1, row),
							CalcUtil.getAttr(sheet, 2, row)
							);
				}
			});
		
		Assert.assertEquals(1, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
	}
	
	@Test
	public void testReadingCalcSpreadsheetWithHolesUsingDefaultErrorHandling() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.ods");		
		CalcTemplate et = new CalcTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow(0, new CalcRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Sheet sheet, int row) {
					PhoneBookEntry entry = new PhoneBookEntry(
							CalcUtil.getAttr(sheet, 0, row),
							CalcUtil.getAttr(sheet, 1, row),
							CalcUtil.getAttr(sheet, 2, row)
							);
					return entry;
				}
			});
		
		Assert.assertEquals(1, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
	}
	
	@Test
	public void testReadingCalcSpreadsheetWithHoles2() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.ods");		
		CalcTemplate et = new CalcTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow(0, new CalcRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Sheet sheet, int row) {
					String name = null;
					try { name = CalcUtil.getAttr(sheet, 0, row); } catch (Exception e) {}
					
					String address = null;
					try { address = CalcUtil.getAttr(sheet, 1, row); } catch (Exception e) {}
					
					String phone = null;
					try { phone = CalcUtil.getAttr(sheet, 2, row); } catch (Exception e) {}
					
					return new PhoneBookEntry(name, address, phone);
				}
			});
		
		Assert.assertEquals(4, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
		
		Assert.assertEquals("Joanna", results.get(1).getName());
		Assert.assertNull(results.get(1).getAddress());
		Assert.assertEquals("555-915-9900", results.get(1).getPhone());
		
		Assert.assertNull(results.get(2).getName());
		Assert.assertEquals("Corp HQ", results.get(2).getAddress());
		Assert.assertEquals("555-321-9502", results.get(2).getPhone());

		Assert.assertEquals("Bill Lumbergh", results.get(3).getName());
		Assert.assertEquals("his cubicle", results.get(3).getAddress());
		Assert.assertNull(results.get(3).getPhone());
	}
	
	@Test
	public void testReadingCalcSpreadsheetWithHolesUsingSpecialErrorHandling() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.ods");		
		CalcTemplate et = new CalcTemplate(file, true);
		
		List<PhoneBookEntry> results = 
			et.onEachRow(0,
					new CalcRowCallback<PhoneBookEntry>() {
						public PhoneBookEntry mapRow(Sheet sheet, int row) {
							return new PhoneBookEntry(
									CalcUtil.getAttr(sheet, 0, row),
									CalcUtil.getAttr(sheet, 1, row),
									CalcUtil.getAttr(sheet, 2, row)
							);
						}
					},
					new CalcTemplateErrorHandler<PhoneBookEntry>() {
						public PhoneBookEntry handleException(Sheet sheet, int row, RuntimeException e) {
							return new EmptyPhoneBookEntry();
						}
					}
			);
		
		Assert.assertEquals(4, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
		
		Assert.assertEquals(EmptyPhoneBookEntry.NAME, results.get(1).getName());
		Assert.assertEquals(EmptyPhoneBookEntry.ADDRESS, results.get(1).getAddress());
		Assert.assertEquals(EmptyPhoneBookEntry.PHONE, results.get(1).getPhone());

		Assert.assertEquals(EmptyPhoneBookEntry.NAME, results.get(2).getName());
		Assert.assertEquals(EmptyPhoneBookEntry.ADDRESS, results.get(2).getAddress());
		Assert.assertEquals(EmptyPhoneBookEntry.PHONE, results.get(2).getPhone());
		
		Assert.assertEquals(EmptyPhoneBookEntry.NAME, results.get(3).getName());
		Assert.assertEquals(EmptyPhoneBookEntry.ADDRESS, results.get(3).getAddress());
		Assert.assertEquals(EmptyPhoneBookEntry.PHONE, results.get(3).getPhone());
	}

}
