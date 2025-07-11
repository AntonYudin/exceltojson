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

import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;

import org.apache.poi.xssf.eventusermodel.XSSFBReader;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;

import org.apache.poi.xssf.usermodel.XSSFComment;


public class XLSBReader implements Reader {

	private final static Logger logger = Logger.getLogger(XLSBReader.class.getName());


	public void read(final InputStream inputStream, final Handler handler) throws Exception {
		logger.info("read(" + inputStream + ", " + handler + ")");

		try (final var pkg = OPCPackage.open(inputStream)) {

			final var reader = new XSSFBReader(pkg);
			final var sst = new XSSFBSharedStringsTable(pkg);
			final var xssfbStylesTable = reader.getXSSFBStylesTable();
			final var iterator = XSSFBReader.SheetIterator.class.cast(reader.getSheetsData());

			while (iterator.hasNext()) {

				final var sheetInputStream = iterator.next();
				final var name = iterator.getSheetName();

				final var contentsHandler = new ContentsHandler(handler);

				handler.startSheet(name);

				final var sheetHandler = new XSSFBSheetHandler(
					sheetInputStream,
					xssfbStylesTable,
					iterator.getXSSFBSheetComments(),
					sst,
					contentsHandler,
					new DataFormatter(),
					false
				);

				sheetHandler.parse();

				contentsHandler.endSheet();
			}
		}
	}


	public static class ContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

		private final Handler handler;


		public ContentsHandler(final Handler handler) {
			this.handler = handler;
		}

		@Override
		public void startRow(final int rowNumber) {
			handler.startRow(rowNumber);
		}

		@Override
		public void endSheet() {
			handler.endSheet();
		}

		@Override
		public void endRow(final int rowNumber) {
			handler.endRow(rowNumber);
		}

		@Override
		public void cell(final String cellReference, final String formattedValue, final XSSFComment comment) {
			handler.startCell(cellReference, formattedValue);
		}

	}

}

