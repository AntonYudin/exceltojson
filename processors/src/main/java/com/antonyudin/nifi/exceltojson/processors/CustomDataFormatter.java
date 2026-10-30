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

import java.text.Format;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.DateUtil;


public class CustomDataFormatter extends DataFormatter {

	private final Format dateFormat;
	private final boolean noFormatting;

	public CustomDataFormatter(final boolean noFormatting, final Format dateFormat) {
		super();
		this.noFormatting = noFormatting;
		this.dateFormat = dateFormat;
		setUseCachedValuesForFormulaCells(false);
	}

	@Override
	public String formatRawCellContents(final double value, final int formatIndex, final String formatString) {
		if (DateUtil.isADateFormat(formatIndex, formatString)) {
			return super.formatRawCellContents(value, formatIndex, formatString);
		}
		if (noFormatting)
			return super.formatRawCellContents(value, 0, "General");
		return super.formatRawCellContents(value, formatIndex, formatString);
	}


	@Override
    protected Format createFormat(final double cellValue, final int formatIndex, final String formatString) {
		if ((dateFormat != null) && DateUtil.isADateFormat(formatIndex, formatString))
			return dateFormat;
		return super.createFormat(cellValue, formatIndex, formatString);
    }

}

