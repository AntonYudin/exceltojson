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


import java.util.logging.Logger;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook ;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.xssf.eventusermodel.XSSFReader;


public class XLSXReader implements Reader {

	private final static Logger logger = Logger.getLogger(XLSXReader.class.getName());


	public void read(final InputStream inputStream, final Handler handler) throws Exception {
		logger.info("read(" + inputStream + ", " + handler + ")");

		try (
			final var pkg = OPCPackage.open(inputStream);
			final var workbook = new XSSFWorkbook(pkg)
		) {

			final var formatter = new DataFormatter();

			final var formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

			final var sheetIterator = workbook.sheetIterator();

			while (sheetIterator.hasNext()) {

				final var sheet = sheetIterator.next();

				handler.startSheet(sheet.getSheetName());

				for (var i = 0; i < sheet.getLastRowNum() + 1; i++) {

					final var row = sheet.getRow(i);

					if (row == null)
						continue;

					handler.startRow(row.getRowNum());

					for (var j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {

						final var cell = row.getCell(j);

						if (cell == null)
							continue;

						final var address = cell.getAddress().toString();

						final var evaluatedCell = formulaEvaluator.evaluateInCell(cell);

						final var value = formatter.formatCellValue(evaluatedCell);

						handler.startCell(address, value);
					}

					handler.endRow(row.getRowNum());
				}

				handler.endSheet();
			}
		}
	}

}

