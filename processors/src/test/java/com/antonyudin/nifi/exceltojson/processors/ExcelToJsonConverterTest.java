/*
 Licensed to the Apache Software Foundation (ASF) under one or more
 contributor license agreements. See the NOTICE file distributed with
 this work for additional information regarding copyright ownership.
 The ASF licenses this file to You under the Apache License, Version 2.0
 (the "License"); you may not use this file except in compliance with
 the License. You may obtain a copy of the License at
 http://www.apache.org/licenses/LICENSE-2.0
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
*/

package com.antonyudin.nifi.exceltojson.processors;


import java.io.ByteArrayOutputStream;

import java.util.logging.Logger;

import org.junit.jupiter.api.Test;


public class ExcelToJsonConverterTest {

	private final static Logger logger = Logger.getLogger(ExcelToJsonConverterTest.class.getName());


	@Test
	public void testXLSX2007() throws Exception {
		try (
				final var inputStream = ExcelToJsonConverterTest.class.getResourceAsStream("/test-excel-2007.xlsx");
				final var outputStream = new ByteArrayOutputStream();
		) {
			System.err.println("inputStream: " + inputStream);
			(new ExcelToJsonConverter()).convert(inputStream, outputStream, null, ExcelToJsonConverter.Type.XLSX, true, true);
			System.err.println("result: " + (new String(outputStream.toByteArray(), "UTF-8")));
		}
	}


	@Test
	public void testXLSXOfficeOpenXMLSpreadsheet() throws Exception {
		try (
				final var inputStream = ExcelToJsonConverterTest.class.getResourceAsStream("/test-office-open-xml-spreadsheet.xlsx");
				final var outputStream = new ByteArrayOutputStream();
		) {
			System.err.println("inputStream: " + inputStream);
			(new ExcelToJsonConverter()).convert(inputStream, outputStream, null, ExcelToJsonConverter.Type.XLSX, true, true);
			System.err.println("result: " + (new String(outputStream.toByteArray(), "UTF-8")));
		}
	}


	@Test
	public void testXLSB() throws Exception {
		try (
				final var inputStream = ExcelToJsonConverterTest.class.getResourceAsStream("/test.xlsb");
				final var outputStream = new ByteArrayOutputStream();
		) {
			System.err.println("inputStream: " + inputStream);
			(new ExcelToJsonConverter()).convert(inputStream, outputStream, null, ExcelToJsonConverter.Type.XLSB, true, true);
			System.err.println("result: " + (new String(outputStream.toByteArray(), "UTF-8")));
		}
	}

}

