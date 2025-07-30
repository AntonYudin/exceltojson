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
import java.io.Serializable;

import jakarta.json.Json;
import jakarta.json.JsonArray;
import jakarta.json.JsonObject;
import jakarta.json.JsonValue;
import jakarta.json.JsonNumber;
import jakarta.json.JsonString;
import jakarta.json.stream.JsonParser;
import jakarta.json.stream.JsonParserFactory;


public class JsonToExcelConverter {

	private final static Logger logger = Logger.getLogger(JsonToExcelConverter.class.getName());


	private final JsonParserFactory jsonFactory = Json.createParserFactory(
		Map.of()
	);


	public enum Type {
		XLSX
	}


	public record Row(Map<String, String> values) {

		public static Row of(final Map<String, JsonValue> values) {
			final var map = new HashMap<String, String>();
			for (var entry: values.entrySet()) {
				map.put(entry.getKey(), entry.getValue().toString());
			}
			return new Row(map);
		}

	}

	public record Sheet(String name, List<Row> rows) {}

	public void convert(
		final InputStream inputStream,
		final OutputStream outputStream,
		final Map<String, String> parameters,
		final Type type,
		final boolean streaming,
		final int autoSizeColumns
	) throws Exception {

		logger.info("convert(" + inputStream + ", " + outputStream + ", " + parameters + ", " + streaming + ", " + autoSizeColumns + ")");

		try (final var writer = new XLSXWriter(outputStream, autoSizeColumns, streaming)) {

			final var parser = jsonFactory.createParser(inputStream);

			parser.next();

			final var sheets = new ArrayList<Sheet>();

			while (parser.hasNext()) {
				final var event = parser.next();
				if (event == JsonParser.Event.KEY_NAME) {
					final var name = parser.getString();
					if (name.equals("sheets")) {
						parser.next();
						sheets.addAll(parseSheets(parser.getArray(), writer));
					}
				}
			}

		}
	}

	protected List<Sheet> parseSheets(final JsonArray sheets, final Writer writer) throws Exception {
		final var result = new ArrayList<Sheet>();
		for (var item: sheets) {
			if (item instanceof JsonObject o) {
				final var name = o.getString("name");
				final var selected = o.getBoolean("selected", false);
				final var active = o.getBoolean("active", false);
				writer.startSheet(name, getStyle(o), selected, active);
				final var rowsArray = o.getJsonArray("rows");
			//	final var rows = new ArrayList<Row>();
				if (rowsArray != null) {
					for (var r: rowsArray) {
						if (r instanceof JsonObject object) {
							writer.addRow();
							for (var entry: object.entrySet()) {
								final var value = entry.getValue();
								addCell(writer, entry.getKey(), value, null);
							}
						}
					}
				}

				{
					final var imagesArray = o.getJsonArray("images");
					if (imagesArray != null) {
						for (var i: imagesArray) {
							if (i instanceof JsonObject object) {
								final var reference = object.getString("reference", null);
								final var url = object.getString("url", null);
								final var type = object.getString("type", null);
								final var scale = object.getJsonNumber("scale");
								writer.addImage(reference, url, Writer.ImageType.valueOf(type), scale == null? 1.0: scale.doubleValue());
							}
						}
					}
				}

				final var printArea = o.getString("printArea", null);
				if (printArea != null)
					writer.setPrintArea(printArea);
				writer.endSheet();
			}
		}
		return result;
	}

	protected Writer.Style getStyle(final JsonObject object) {

		Writer.Style result = null;

		final var alignment = object.getString("alignment", null);
		if (alignment != null)
			result = Writer.Style.aligned(Writer.Alignment.valueOf(alignment), result);

		final var fontHeight = object.getInt("fontHeight", -1);
		if (fontHeight >= 0)
			result = Writer.Style.fontHeight(fontHeight, result);

		final var color = object.getString("color", null);
		if (color != null)
			result = Writer.Style.color(color, result);

		final var fillColor = object.getString("fillColor", null);
		if (fillColor != null)
			result = Writer.Style.fillColor(fillColor, result);

		if (object.getBoolean("fontWeightBold", false))
			result = Writer.Style.fontWeightBold(true, result);

		final var width = object.getInt("width", -1);
		if (width >= 0)
			result = Writer.Style.width(width, result);

		return result;
	}

	protected void addCell(final Writer writer, final String name, final JsonValue value, final Writer.Style style) {

		if (value == null) {
			writer.addColumn(name, (String) null, style);
			return;
		}

		switch (value) {
			case JsonObject object -> {

				Writer.Style newStyle = getStyle(object);

				if (object.getBoolean("formula", false)) {
					writer.addColumn(name, object.getString("value"), true, newStyle);
				} else {
					addCell(writer, name, object.get("value"), newStyle);
				}
				final var columns = object.getInt("columns", 0);

				if (columns > 1)
					writer.mergeColumns(columns);
			}
			case JsonNumber number -> writer.addColumn(name, number.bigDecimalValue(), style);
			case JsonString string -> writer.addColumn(name, string.getString(), style);
			default -> {
				if (value == JsonValue.TRUE)
					writer.addColumn(name, true, style);
				else if (value == JsonValue.FALSE)
					writer.addColumn(name, false, style);
				else if (value == JsonValue.NULL)
					writer.addColumn(name, (String) null, style);
			}
		}
	}

}

