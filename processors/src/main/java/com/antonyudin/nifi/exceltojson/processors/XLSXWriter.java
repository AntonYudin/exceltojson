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

import java.net.URL;

import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;
import java.util.HexFormat;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.ClientAnchor;

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
	private final Style defaultStyle = Style.of();
	private CellStyle defaultRowStyle = null;


	public XLSXWriter(final OutputStream outputStream, final int autoSizeColumns, final boolean streaming) {
	//	workbook = (streaming? new SXSSFWorkbook(null, 100, false, false): new XSSFWorkbook());
		workbook = (streaming? new SXSSFWorkbook(): new XSSFWorkbook());
		if (workbook instanceof SXSSFWorkbook w) {
		//	w.setZip64Mode(Zip64Mode.AlwaysWithCompatibility);
			w.setZip64Mode(Zip64Mode.AsNeeded);
		}
		this.outputStream = outputStream;
		this.autoSizeColumns = autoSizeColumns;
		defaultRowStyle = workbook.createCellStyle();
	//	defaultRowStyle.setWrapText(true);
	}

	@Override
	public void close() throws Exception {
		workbook.write(outputStream);
		workbook.close();
		switch (workbook) {
			case SXSSFWorkbook sxssfWorkbook -> sxssfWorkbook.dispose();
			default -> {}
		}
	}

	@Override
	public void startSheet(final String name, final Style style, final boolean selected, final boolean active) {
		images.clear();
		sheet = workbook.createSheet(name);
		sheetIndex++;
		if (autoSizeColumns != 0) {
			if (sheet instanceof SXSSFSheet s)
				s.trackAllColumnsForAutoSizing();
		}
		rowNumber = 0;
		row = null;
		maxFontHeight = 0;
		maxCellNumber = -1;
		autoSized = false;
		if (style != null) {
			if (style.color() != null) {
				switch (sheet) {
					case XSSFSheet xssfSheet -> xssfSheet.setTabColor(getColor(style.color()));
					case SXSSFSheet sxssfSheet -> sxssfSheet.setTabColor(getColor(style.color()));
					default -> throw new IllegalArgumentException("cannot set color for [" + sheet + "]"); 
				}
			}
		}
		if (selected)
			sheet.setSelected(selected);
		if (active)
			workbook.setActiveSheet(sheetIndex);
	}

	protected void setHeight() {
		if (row == null)
			return;
		if (rowNumber >= 0) {
			if (maxFontHeight > 0) {
				row.setHeightInPoints((short) (maxFontHeight + 6));
			} else
				row.setHeight((short) -1);
		}
	}

	@Override
	public void endSheet() throws Exception {

		setHeight();

		if ((autoSizeColumns != 0) && (!autoSized)) {
			for (var i = 0; i <= maxCellNumber; i++)
				sheet.autoSizeColumn(i, true);
		}

		addImages();

		if (sheet instanceof SXSSFSheet s) {
			s.flushRows();
			s.validateMergedRegions();
		}

	}

	private boolean autoSized = false;

	@Override
	public void addRow() {

		setHeight();

		maxFontHeight = 0;

		row = sheet.createRow(rowNumber++);
		
		switch (row) {
			case XSSFRow xssfRow -> xssfRow.setRowStyle(defaultRowStyle);
			case SXSSFRow sxssfRow -> sxssfRow.setRowStyle(defaultRowStyle);
			default -> {}
		}

		row.setHeight((short) -1);

		cellNumber = -1;

		if ((rowNumber % 10000) == 0)
			logger.info("wrote [" + rowNumber + "] rows.");

		if (autoSizeColumns > 0) {

			if (!autoSized) {

				if (rowNumber > autoSizeColumns) {

					logger.info("autosizing ...");

					for (var i = 0; i <= maxCellNumber; i++)
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

	private final static HexFormat hexFormat = HexFormat.of().withDelimiter("").withPrefix("").withSuffix("");

	protected XSSFColor getColor(final String color) {
		final var result = new XSSFColor(
			switch (workbook) {
				case XSSFWorkbook xssf -> xssf.getStylesSource().getIndexedColors();
				case SXSSFWorkbook sxssf -> sxssf.getXSSFWorkbook().getStylesSource().getIndexedColors();
				default -> throw new IllegalArgumentException("unsupported workbook: " + workbook);
			}
		);
		result.setARGBHex(color);
		return result;
	}

	private int maxFontHeight = 0;

	protected void setStyle(final Cell cell, final Style style) {
		final var effectiveStyle = (style == null? defaultStyle: style);
		var s = stylesCache.get(effectiveStyle);
		if (s == null) {
			s = workbook.createCellStyle();
			//s.setWrapText(true);
			if (effectiveStyle.alignment() != null) {
				switch (effectiveStyle.alignment()) {
					case Alignment.center: s.setAlignment(HorizontalAlignment.CENTER); break;
					case Alignment.left: s.setAlignment(HorizontalAlignment.LEFT); break;
					case Alignment.right: s.setAlignment(HorizontalAlignment.RIGHT); break;
				}
			}

			if (
				(effectiveStyle.fontHeight() != null) ||
				(effectiveStyle.fontWeightBold() != null) ||
				(effectiveStyle.color() != null)
			) {

				final var font = workbook.createFont();

				if (effectiveStyle.fontHeight() != null)
					font.setFontHeightInPoints(effectiveStyle.fontHeight().shortValue());

				if (effectiveStyle.fontWeightBold() != null)
					font.setBold(effectiveStyle.fontWeightBold());
				if (effectiveStyle.color() != null) {
					switch (font) {
						case XSSFFont xssfFont -> {
							final var color = getColor(effectiveStyle.color());
							xssfFont.setColor(color);
						}
						default -> throw new IllegalArgumentException("unsupported font: " + font);
					}
				}

				s.setFont(font);
			}

			if (effectiveStyle.fillColor() != null) {
				s.setFillForegroundColor(getColor(effectiveStyle.fillColor()));
				s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			}

			stylesCache.put(effectiveStyle, s);
		} else {
			// logger.info("reusing style: [" + style + "]");
		}

		if (effectiveStyle.fontHeight() != null) {
			final var height = effectiveStyle.fontHeight();
			if (height > maxFontHeight)
				maxFontHeight =  height;
		}

		cell.setCellStyle(s);
	}

	@Override
	public void mergeColumns(final int numberOfColumns) {
		sheet.addMergedRegion(new CellRangeAddress(rowNumber - 1, rowNumber - 1, cellNumber, cellNumber + numberOfColumns - 1));
		cellNumber += (numberOfColumns - 1);
	}

	@Override
	public void setPrintArea(final String reference) {
		workbook.setPrintArea(sheetIndex, reference);
	}


	public record Image(String reference, String url, ImageType type, double scale) {}

	private final List<Image> images = new ArrayList<>();

	@Override
	public void addImage(final String reference, final String imageURL, final ImageType type, final double scale) throws Exception {
		images.add(new Image(reference, imageURL, type, scale));
	}

	protected void addImages() throws Exception {

		for (var image: images) {

			final var url = new URL(image.url());

			try (final var stream = url.openStream()) {

				final var index = workbook.addPicture(stream.readAllBytes(), switch (image.type()) {
					case ImageType.png -> workbook.PICTURE_TYPE_PNG;
					case ImageType.jpeg -> workbook.PICTURE_TYPE_JPEG;
				});

				final var drawing = sheet.createDrawingPatriarch();
				final var helper = workbook.getCreationHelper();
				final var anchor = helper.createClientAnchor();

				final var cellRange = CellRangeAddress.valueOf(image.reference());

				//anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
				//anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
				//anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);

				anchor.setCol1(cellRange.getFirstColumn());
			//	anchor.setCol2(cellRange.getLastColumn());
				anchor.setCol2(-1);
				anchor.setRow1(cellRange.getFirstRow());
			//	anchor.setRow2(cellRange.getLastRow());
				anchor.setRow2(-1);

				final var picture = drawing.createPicture(anchor, index);

				picture.resize(image.scale());
			}
		}
	}

}

