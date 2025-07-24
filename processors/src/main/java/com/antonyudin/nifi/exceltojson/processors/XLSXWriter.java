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

import java.io.OutputStream;

import java.math.BigDecimal;

import java.util.Map;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import org.apache.commons.compress.archivers.zip.Zip64Mode;

import org.apache.poi.ss.util.CellRangeAddress;



public class XLSXWriter implements Writer, AutoCloseable {

	private final static Logger logger = Logger.getLogger(XLSXWriter.class.getName());

	private final OutputStream outputStream;
	private final Workbook workbook;
	private final int autoSizeColumns;
	private Sheet sheet;
	private Row row;
	private int sheetIndex = -1;
	private int rowNumber;
	private int cellNumber;
	private int maxCellNumber;


	public XLSXWriter(final OutputStream outputStream, final int autoSizeColumns, final boolean streaming) {
		workbook = (streaming? new SXSSFWorkbook(): new XSSFWorkbook());
		if (workbook instanceof SXSSFWorkbook w) {
			w.setZip64Mode(Zip64Mode.AlwaysWithCompatibility);
		}
		this.outputStream = outputStream;
		this.autoSizeColumns = autoSizeColumns;
	}

	@Override
	public void close() throws Exception {
		workbook.write(outputStream);
	}

	@Override
	public void startSheet(final String name) {
		sheet = workbook.createSheet(name);
		sheetIndex++;
		if (autoSizeColumns != 0) {
			if (sheet instanceof SXSSFSheet s)
				s.trackAllColumnsForAutoSizing();
		}
		rowNumber = 0;
		maxCellNumber = -1;
		autoSized = false;
	}

	@Override
	public void endSheet() {
		if ((autoSizeColumns != 0) && (!autoSized)) {
			for (var i = 0; i < maxCellNumber; i++)
				sheet.autoSizeColumn(i, true);
		}
	}

	private boolean autoSized = false;

	@Override
	public void addRow() {

		row = sheet.createRow(rowNumber++);

		cellNumber = -1;

		if ((rowNumber % 10000) == 0)
			logger.info("wrote [" + rowNumber + "] rows.");

		if (autoSizeColumns > 0) {

			if (!autoSized) {

				if (rowNumber > autoSizeColumns) {

					logger.info("autosizing ...");

					for (var i = 0; i < maxCellNumber; i++)
						sheet.autoSizeColumn(i, true);

					if (sheet instanceof SXSSFSheet s)
						s.untrackAllColumnsForAutoSizing();

					autoSized = true;

					logger.info("autosizing ... done.");
				}
			}
		}
	}

	protected int nextCell() {
		cellNumber++;
		if (cellNumber > maxCellNumber)
			maxCellNumber = cellNumber;
		return cellNumber;
	}

	@Override
	public void addColumn(final String name, final String value, final Style style) {
		final var cell = row.createCell(nextCell());
		cell.setCellValue(value);
		setStyle(cell, style);
	}

	@Override
	public void addColumn(final String name, final boolean value, final Style style) {
		final var cell = row.createCell(nextCell());
		cell.setCellValue(value);
		setStyle(cell, style);
	}

	@Override
	public void addColumn(final String name, final int value, final Style style) {
		final var cell = row.createCell(nextCell());
		cell.setCellValue(value);
		setStyle(cell, style);
	}

	@Override
	public void addColumn(final String name, final BigDecimal value, final Style style) {
		final var cell = row.createCell(nextCell());
		cell.setCellValue(value.doubleValue());
		setStyle(cell, style);
	}

	@Override
	public void addColumn(final String name, final String value, final boolean formula, final Style style) {
		if (!formula) {
			addColumn(name, value, style);
			return;
		}
		final var cell = row.createCell(nextCell());
		cell.setCellFormula(value);
		setStyle(cell, style);
	}

	private Map<Style, CellStyle> stylesCache = new HashMap<>();

	protected void setStyle(final Cell cell, final Style style) {
		if (style == null)
			return;
		var s = stylesCache.get(style);
		if (s == null) {
			s = workbook.createCellStyle();
			if (style.alignment() != null) {
				switch (style.alignment()) {
					case Alignment.center: s.setAlignment(HorizontalAlignment.CENTER); break;
					case Alignment.left: s.setAlignment(HorizontalAlignment.LEFT); break;
					case Alignment.right: s.setAlignment(HorizontalAlignment.RIGHT); break;
				}
			}
			if (style.fontHeight() != null) {
				final var font = workbook.createFont();
				font.setFontHeightInPoints(style.fontHeight().shortValue());
				s.setFont(font);
			}
			stylesCache.put(style, s);
		} else {
			// logger.info("reusing style: [" + style + "]");
		}
		cell.setCellStyle(s);
	}

	@Override
	public void mergeColumns(final int numberOfColumns) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber - 1, rowNumber - 1, cellNumber, cellNumber + numberOfColumns - 1));
	}

	@Override
	public void setPrintArea(final String reference) {
		workbook.setPrintArea(sheetIndex, reference);
	}

}

