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
import java.util.List;

import junit.framework.Assert;

import org.apache.poi.ss.usermodel.Row;
import org.junit.Test;
import org.springframework.batch.spreadsheet.support.EmptyPhoneBookEntry;
import org.springframework.batch.spreadsheet.support.PhoneBookEntry;

/**
 * @author Greg Turnquist
 */
public class TestExcelTemplate {

	private String pathname = "src" + File.separator + "test" + File.separator + "resources";
	
	@Test
	public void testReadingSimpleExcelSpreadsheetWithoutHeaderHandling() {
		File file = new File(pathname + File.separator + "phonebook.xls");		
		ExcelTemplate et = new ExcelTemplate(file);
		List<PhoneBookEntry> results = 
			et.onEachRow("Sheet1", new ExcelRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Row row) {
					return new PhoneBookEntry(
							row.getCell(0).getStringCellValue(),
							row.getCell(1).getStringCellValue(),
							row.getCell(2).getStringCellValue()
							);
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
	public void testReadingSimpleExcelSpreadsheetSkippingHeader() {
		File file = new File(pathname + File.separator + "phonebook.xls");		
		ExcelTemplate et = new ExcelTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow("Sheet1", new ExcelRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Row row) {
					return new PhoneBookEntry(
							row.getCell(0).getStringCellValue(),
							row.getCell(1).getStringCellValue(),
							row.getCell(2).getStringCellValue()
							);
				}
			});
		
		Assert.assertEquals(1, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
	}
	
	@Test
	public void testReadingExcelSpreadsheetWithHolesUsingDefaultErrorHandling() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.xls");		
		ExcelTemplate et = new ExcelTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow("Sheet1", new ExcelRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Row row) {
					return new PhoneBookEntry(
							row.getCell(0).getStringCellValue(),
							row.getCell(1).getStringCellValue(),
							row.getCell(2).getStringCellValue()
							);
				}
			});
		
		Assert.assertEquals(1, results.size());
		
		Assert.assertEquals("Peter Gibbons", results.get(0).getName());
		Assert.assertEquals("123 ABC Drive", results.get(0).getAddress());
		Assert.assertEquals("555-821-2123", results.get(0).getPhone());
	}
	
	@Test
	public void testReadingExcelSpreadsheetWithHoles2() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.xls");		
		ExcelTemplate et = new ExcelTemplate(file, true);
		List<PhoneBookEntry> results = 
			et.onEachRow("Sheet1", new ExcelRowCallback<PhoneBookEntry>() {
				public PhoneBookEntry mapRow(Row row) {
					String name = null;
					try { name = row.getCell(0).getStringCellValue(); } catch (Exception e) {}
					
					String address = null;
					try { address = row.getCell(1).getStringCellValue(); } catch (Exception e) {}
					
					String phone = null;
					try { phone = row.getCell(2).getStringCellValue(); } catch (Exception e) {}
					
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
	public void testReadingExcelSpreadsheetWithHolesUsingSpecialErrorHandling() {
		File file = new File(pathname + File.separator + "phonebook_with_holes.xls");		
		ExcelTemplate et = new ExcelTemplate(file, true);
		
		List<PhoneBookEntry> results = 
			et.onEachRow("Sheet1",
					new ExcelRowCallback<PhoneBookEntry>() {
						public PhoneBookEntry mapRow(Row row) {
							return new PhoneBookEntry(
									row.getCell(0).getStringCellValue(),
									row.getCell(1).getStringCellValue(),
									row.getCell(2).getStringCellValue()
							);
						}
					},
					new ExcelTemplateErrorHandler<PhoneBookEntry>() {
						public PhoneBookEntry handleException(Row row, RuntimeException e) {
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
