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

import java.io.ByteArrayOutputStream;

import org.apache.nifi.components.PropertyDescriptor;
import org.apache.nifi.components.PropertyValue;

import org.apache.nifi.flowfile.FlowFile;

import org.apache.nifi.annotation.behavior.ReadsAttribute;
import org.apache.nifi.annotation.behavior.ReadsAttributes;
import org.apache.nifi.annotation.behavior.WritesAttribute;
import org.apache.nifi.annotation.behavior.WritesAttributes;

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

import org.apache.nifi.flowfile.attributes.CoreAttributes;


@Tags({"xlsb", "xlsx", "excel", "json"})
@CapabilityDescription("Transforms XLSB/XLSX files to Json")
@SeeAlso({})
@ReadsAttributes({@ReadsAttribute(attribute="", description="")})
@WritesAttributes({@WritesAttribute(attribute="", description="")})
public class ExcelToJson extends AbstractProcessor {

	private final static Logger logger = Logger.getLogger(ExcelToJson.class.getName());

	public static final PropertyDescriptor AUTODETECT_TYPES_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("Autodetect Types")
		.displayName("Autodetect Types")
		.description("Autodetect Types")
		.required(true)
		.allowableValues("true", "false")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;

	public static final PropertyDescriptor PRETTY_PRINTING_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("Pretty Printing")
		.displayName("Pretty Printing")
		.description("Pretty Printing")
		.required(true)
		.allowableValues("true", "false")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;

	public static final PropertyDescriptor FILE_TYPE_PROPERTY = new PropertyDescriptor
		.Builder()
		.name("File Type")
		.displayName("File Type")
		.description("File Type")
		.required(true)
		.allowableValues("xlsb", "xlsx", "use file extension")
		.addValidator(StandardValidators.NON_EMPTY_VALIDATOR)
		.build()
	;


	public static final Relationship REL_SUCCESS = new Relationship.Builder()
		.name("success")
		.description("JSON success relationship")
		.build()
	;

	private List<PropertyDescriptor> descriptors;

	private Set<Relationship> relationships;

	@Override
	protected void init(final ProcessorInitializationContext context) {
		descriptors = List.of(FILE_TYPE_PROPERTY, AUTODETECT_TYPES_PROPERTY, PRETTY_PRINTING_PROPERTY);
		relationships = Set.of(REL_SUCCESS);
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

		final var buffer = new ByteArrayOutputStream();

		final var mappedProperties = new HashMap<String, String>();

		for (var entry: context.getProperties().entrySet()) {
			if (entry.getKey().isDynamic())
				mappedProperties.put(entry.getKey().getName(), entry.getValue());
		}

		final var autodetectTypes = context.getProperty(AUTODETECT_TYPES_PROPERTY.getName());
		final var prettyPrinting = context.getProperty(PRETTY_PRINTING_PROPERTY.getName());
		final var fileType = context.getProperty(FILE_TYPE_PROPERTY.getName());

		final var fileName = flowFile.getAttribute(CoreAttributes.FILENAME.key());


		session.read(flowFile, (inputStream) -> {
			try {
				(new ExcelToJsonConverter()).convert(
					inputStream,
					buffer,
					mappedProperties,
					parseFileType(fileType, fileName),
					autodetectTypes != null && autodetectTypes.getValue().equalsIgnoreCase("true"),
					prettyPrinting != null && prettyPrinting.getValue().equalsIgnoreCase("true")
				);
		} catch (Exception exception) {
				logger.log(Level.SEVERE, "exception", exception);
			}
		});

		session.write(flowFile, (outputStream) -> {
			buffer.writeTo(outputStream);
		});

		session.transfer(flowFile, REL_SUCCESS);
	}

	protected ExcelToJsonConverter.Type parseFileType(final PropertyValue propertyValue, final String fileName) {

		if (propertyValue == null)
			return ExcelToJsonConverter.Type.XLSB;

		final var value = propertyValue.getValue();

		if (value.equalsIgnoreCase("xlsb"))
			return ExcelToJsonConverter.Type.XLSB;

		if (value.equalsIgnoreCase("xlsx"))
			return ExcelToJsonConverter.Type.XLSX;

		if (value.equalsIgnoreCase("use file extension")) {
			final var lower = fileName.toLowerCase();
			if (lower.endsWith(".xlsx"))
				return ExcelToJsonConverter.Type.XLSX;
			if (lower.endsWith(".xlsb"))
				return ExcelToJsonConverter.Type.XLSB;
			return null;
		}

		return null;
	}

}

