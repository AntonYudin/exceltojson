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
import java.util.logging.Level;

import java.util.List;
import java.util.Set;
import java.util.HashMap;

import java.io.StringWriter;
import java.io.PrintWriter;
import java.io.ByteArrayOutputStream;

import org.apache.nifi.components.PropertyDescriptor;
import org.apache.nifi.components.PropertyValue;

import org.apache.nifi.flowfile.FlowFile;

import org.apache.nifi.annotation.behavior.ReadsAttribute;
import org.apache.nifi.annotation.behavior.ReadsAttributes;
import org.apache.nifi.annotation.behavior.WritesAttribute;
import org.apache.nifi.annotation.behavior.WritesAttributes;

import org.apache.nifi.serialization.record.Record;
import org.apache.nifi.serialization.record.RecordSchema;

import org.apache.nifi.annotation.documentation.CapabilityDescription;
import org.apache.nifi.annotation.documentation.SeeAlso;
import org.apache.nifi.annotation.documentation.Tags;

import org.apache.nifi.annotation.lifecycle.OnScheduled;

import org.apache.nifi.processor.AbstractProcessor;

import org.apache.nifi.processor.ProcessContext;
import org.apache.nifi.processor.ProcessSession;
import org.apache.nifi.processor.ProcessorInitializationContext;

import org.apache.nifi.processor.Relationship;

import org.apache.nifi.processor.util.StandardValidators;

import org.apache.nifi.processors.standard.AbstractRecordProcessor;

import org.apache.nifi.flowfile.attributes.CoreAttributes;


@Tags({"xlsx", "excel", "json"})
@CapabilityDescription("Transforms Json to XLSX")
@SeeAlso({})
@ReadsAttributes({@ReadsAttribute(attribute="", description="")})
@WritesAttributes({@WritesAttribute(attribute="", description="")})
public class JsonToExcel extends AbstractProcessor {

	private final static Logger logger = Logger.getLogger(JsonToExcel.class.getName());


	public static final PropertyDescriptor STREAMING_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("Use Streaming POI Workbook")
		.displayName("Use Streaming POI Workbook")
		.description("Use Streaming POI Workbook")
		.required(true)
		.allowableValues("true", "false")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;

	public static final PropertyDescriptor AUTOSIZECOLUMNS_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("Auto Size Columns by tracking")
		.displayName("Auto Size Columns by tracking")
		.description("-1: track all columns; 0: do not track columns; number: how many rows to track")
		.required(true)
	//	.allowableValues("true", "false")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;

	public static final PropertyDescriptor FILE_TYPE_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("File Type")
		.displayName("File Type")
		.description("File Type")
		.required(true)
		.allowableValues("xlsx", "use file extension")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;


	public static final Relationship REL_SUCCESS = new Relationship.Builder()
		.name("success")
		.description("success relationship")
		.build()
	;

	public static final Relationship REL_FAILURE = new Relationship.Builder()
		.name("failure")
		.description("failure relationship")
		.build()
	;

	private List<PropertyDescriptor> descriptors;

	private Set<Relationship> relationships;

	@Override
	protected void init(final ProcessorInitializationContext context) {
		descriptors = List.of(FILE_TYPE_PROPERTY, STREAMING_PROPERTY, AUTOSIZECOLUMNS_PROPERTY);
		relationships = Set.of(REL_SUCCESS, REL_FAILURE);
	}

	@Override
	public Set<Relationship> getRelationships() {
		return this.relationships;
	}

	@Override
	public final List<PropertyDescriptor> getSupportedPropertyDescriptors() {
		return descriptors;
	}

	@Override
	public PropertyDescriptor getSupportedDynamicPropertyDescriptor(final String name) {
		return new PropertyDescriptor.Builder()
			.name(name)
			.displayName(name)
			.description(name)
			.required(false)
			.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
			.dynamic(true)
			.build()
		;
	}


	@OnScheduled
	public void onScheduled(final ProcessContext context) {
	}

	@Override
	public void onTrigger(final ProcessContext context, final ProcessSession session) {

		final var flowFile = session.get();

		if (flowFile == null)
			return;

		try {

			final var buffer = new ByteArrayOutputStream();

			final var mappedProperties = new HashMap<String, String>();

			for (var entry: context.getProperties().entrySet()) {
				if (entry.getKey().isDynamic())
					mappedProperties.put(entry.getKey().getName(), entry.getValue());
			}

			final var streaming = context.getProperty(STREAMING_PROPERTY.getName());
			final var autoSizeColumns = context.getProperty(AUTOSIZECOLUMNS_PROPERTY.getName());
			final var fileType = context.getProperty(FILE_TYPE_PROPERTY.getName());

			final var fileName = flowFile.getAttribute(CoreAttributes.FILENAME.key());


			try (final var inputStream = session.read(flowFile)) {
				(new JsonToExcelConverter()).convert(
					inputStream,
					buffer,
					mappedProperties,
					parseFileType(fileType, fileName),
					streaming != null && streaming.getValue().equalsIgnoreCase("true"),
					autoSizeColumns != null? Integer.valueOf(autoSizeColumns.getValue()): 0
				);
			}

			session.write(flowFile, (outputStream) -> {
				buffer.writeTo(outputStream);
			});

			session.transfer(flowFile, REL_SUCCESS);

		} catch (Exception exception) {
			final var attributes = new HashMap<String, String>();
			attributes.put("exception", exception.getMessage());
			try (final var stringWriter = new StringWriter(); final var printWriter = new PrintWriter(stringWriter)) {
				exception.printStackTrace(printWriter);
				attributes.put("stackTrace", stringWriter.toString());
			} catch (Exception e) {
				logger.log(Level.SEVERE, "exception", e);
			}
			session.putAllAttributes(flowFile, attributes);
			session.transfer(flowFile, REL_FAILURE);
		}
	}

	protected JsonToExcelConverter.Type parseFileType(final PropertyValue propertyValue, final String fileName) {

		if (propertyValue == null)
			return JsonToExcelConverter.Type.XLSX;

		final var value = propertyValue.getValue();

		if (value.equalsIgnoreCase("xlsx"))
			return JsonToExcelConverter.Type.XLSX;

		if (value.equalsIgnoreCase("use file extension")) {
			final var lower = fileName.toLowerCase();
			if (lower.endsWith(".xlsx"))
				return JsonToExcelConverter.Type.XLSX;
			return null;
		}

		return null;
	}

}

