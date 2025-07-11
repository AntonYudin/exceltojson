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

import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;

import java.io.InputStream;
import java.io.OutputStream;

import jakarta.json.Json;
import jakarta.json.stream.JsonGenerator;
import jakarta.json.stream.JsonGeneratorFactory;


public class ExcelToJsonConverter {

	private final static Logger logger = Logger.getLogger(ExcelToJsonConverter.class.getName());


	private final JsonGeneratorFactory jsonFactoryDefault = Json.createGeneratorFactory(
		Map.of()
	);

	private final JsonGeneratorFactory jsonFactoryPrettyPrinting = Json.createGeneratorFactory(
		Map.of(JsonGenerator.PRETTY_PRINTING, true)
	);


	public enum Type {
		XLSX,
		XLSB
	}

	public void convert(
		final InputStream inputStream,
		final OutputStream outputStream,
		final Map<String, String> parameters,
		final Type type,
		final boolean autodetectTypes,
		final boolean prettyPrinting
	) throws Exception {

		logger.info("convert(" + inputStream + ", " + outputStream + ", " + parameters + ", " + autodetectTypes + ", " + prettyPrinting + ")");

		final var json = (
			prettyPrinting?
			jsonFactoryPrettyPrinting.createGenerator(outputStream):
			jsonFactoryDefault.createGenerator(outputStream)
		);

		json.writeStartObject();
		json.writeStartArray("sheets");

		final var contentHandler = new ContentsHandler(
			json,
			(sheetName, column) -> {
				final var renamed = (parameters != null? parameters.get(sheetName + "." + column): null);
				return (renamed != null? renamed: column);
			},
			(sheetName) -> parseLong(parameters != null? parameters.get(sheetName + ".headerRows"): null, 0),
			(sheetName) -> parseBoolean(parameters != null? parameters.get(sheetName + ".autoColumns"): null, false),
			autodetectTypes
		);

		switch (type) {
			case Type.XLSX: (new XLSXReader()).read(inputStream, contentHandler); break;
			case Type.XLSB: (new XLSBReader()).read(inputStream, contentHandler); break;
			default:
				throw new IllegalArgumentException("unknown type: [" + type + "]");
		}

		json.writeEnd();
		json.writeEnd();
		json.close();
	}

	protected static long parseLong(final String value, final long defaultValue) {
		if (value == null)
			return defaultValue;
		return Long.valueOf(value);
	}

	protected static boolean parseBoolean(final String value, final boolean defaultValue) {
		if (value == null)
			return defaultValue;
		return Boolean.valueOf(value);
	}

	public static class ContentsHandler implements Handler {

		public interface ColumnMapper {
			public String map(final String sheetName, final String name);
		}

		public interface HeaderRowsGetter {
			public long getNumberOfHeaderRows(final String sheetName);
		}

		public interface AutoColumnsEnabledGetter {
			public boolean isAutoColumnsEnabled(final String sheetName);
		}

		private final ColumnMapper mapper;
		private final HeaderRowsGetter headerRowsGetter;
		private final AutoColumnsEnabledGetter autoColumnsEnabledGetter;
		private final boolean autodetectTypes;
		private final Map<String, List<String>> headerColumnNames = new HashMap<>();
		private final JsonGenerator json;

		private long headerRows = -1;
		private boolean autoColumns = false;
		private boolean headerRendered = false;
		private boolean processingHeader = true;
		private boolean contentRendered = false;
		private int lastRowNumber = -1;


		public ContentsHandler(
			final JsonGenerator json,
			final ColumnMapper mapper,
			final HeaderRowsGetter headerRowsGetter,
			final AutoColumnsEnabledGetter autoColumnsEnabledGetter,
			final boolean autodetectTypes
		) {
			this.json = json;
			this.mapper = mapper;
			this.headerRowsGetter = headerRowsGetter;
			this.autoColumnsEnabledGetter = autoColumnsEnabledGetter;
			this.autodetectTypes = autodetectTypes;
		}

		private String sheetName;

		@Override
		public void startSheet(final String name) {
			sheetName = name;
			headerRendered = false;
			contentRendered = false;
			processingHeader = true;
			lastRowNumber = -1;

			headerRows = headerRowsGetter.getNumberOfHeaderRows(sheetName);
			autoColumns = autoColumnsEnabledGetter.isAutoColumnsEnabled(sheetName);

			json.writeStartObject();
			json.write("name", name);
		}

		@Override
		public void startRow(final int rowNumber) {
			if ((headerRows > 0) && (!headerRendered)) {
				json.writeStartObject("header");
				json.writeStartArray("rows");
				headerRendered = true;
			}
			if (((headerRows <= 0) || (rowNumber >= headerRows)) && (!contentRendered)) {
				if (headerRendered) {
					json.writeEnd();
					json.writeEnd();
				}
				json.writeStartObject("content");
				json.writeStartArray("rows");
				contentRendered = true;
			}
			if (lastRowNumber < 0) {
				for (var i = 0; i < rowNumber; i++) {
					json.writeStartObject();
					json.writeEnd();
				}
			}
			json.writeStartObject();
			lastRowNumber = rowNumber;
		}

		@Override
		public void endSheet() {
			json.writeEnd(); // rows
			json.writeEnd(); // content
			json.writeEnd(); // sheet
		}

		@Override
		public void endRow(final int rowNumber) {
			json.writeEnd();
		}

		public interface Type {
			public void write(final String name, final JsonGenerator json, final String value) throws Exception;
		}

		private final static Type[] types = {
			(name, json, value) -> json.write(name, Long.valueOf(value)),
			(name, json, value) -> json.write(name, Double.valueOf(value)),
			(name, json, value) -> {
				if (value.equals("true") || value.equals("false"))
					json.write(name, Boolean.valueOf(value));
				throw new IllegalArgumentException("not boolean");
			},
			(name, json, value) -> json.write(name, value)
		};

		@Override
		public void startCell(final String cellReference, final String formattedValue) {
			final var columnReference = getColumnReference(cellReference);
			final var columnName = getColumnName(cellReference);
			if (!contentRendered) {
				var list = headerColumnNames.get(columnReference);
				if (list == null) {
					list = new ArrayList<>();
					headerColumnNames.put(columnReference, list);
				}
				list.add(formattedValue);
			}
			if (autodetectTypes) {
				for (var type: types) {
					try {
						type.write(columnName, json, formattedValue);
						break;
					} catch (Exception e0) {
					}
				}
			} else {
				json.write(columnName, formattedValue);
			}
		}

		protected String getColumnName(final String reference) {
			final var columnReference = getColumnReference(reference);
			if (columnReference == null)
				return null;
			if (!contentRendered)
				return columnReference;
			if (autoColumns) {
				final var list = headerColumnNames.get(columnReference);
				if ((list != null) && (list.size() > 0))
					return String.join(" - ", list);
			}
			return mapper.map(sheetName, columnReference);
		}

		protected String getColumnReference(final String reference) {
			if (reference == null)
				return null;
			for (var i = 0; i < reference.length(); i++) {
				if (!Character.isAlphabetic(reference.charAt(i)))
					return reference.substring(0, i);
			}
			return reference;
		}

	}

}

