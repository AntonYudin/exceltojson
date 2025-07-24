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
import java.io.FileOutputStream;

import java.util.logging.Logger;

import org.junit.jupiter.api.Test;


public class JsonToExcelConverterTest {

	private final static Logger logger = Logger.getLogger(JsonToExcelConverterTest.class.getName());


	@Test
	public void test() throws Exception {

		try (
				final var inputStream = JsonToExcelConverterTest.class.getResourceAsStream("/test.json");
				final var outputStream = new FileOutputStream("test.xlsx");
		) {
			System.err.println("inputStream: " + inputStream);
			(new JsonToExcelConverter()).convert(inputStream, outputStream, null, JsonToExcelConverter.Type.XLSX, true, 4);
//			System.err.println("result: " + (new String(outputStream.toByteArray(), "UTF-8")));
		}
	}

}

