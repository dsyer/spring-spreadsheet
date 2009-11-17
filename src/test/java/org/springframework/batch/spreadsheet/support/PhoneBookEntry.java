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

package org.springframework.batch.spreadsheet.support;

import org.springframework.batch.spreadsheet.CalcTemplate;
import org.springframework.batch.spreadsheet.ExcelTemplate;

/**
 * Simple POJO to represent an empty phone book entry.
 * 
 * @since 11/2/2009
 * @author Greg Turnquist
 * @see CalcTemplate
 * @see ExcelTemplate
 */
public class PhoneBookEntry {
	
	private String name;
	private String address;
	private String phone;

	public PhoneBookEntry() {
	}
	
	public PhoneBookEntry(String name, String address, String phone) {
		this.name = name;
		this.address = address;
		this.phone = phone;
	}
	
	public String toString() {
		return "Name = " + name + " Address = " + address + " Phone = " + phone;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

}