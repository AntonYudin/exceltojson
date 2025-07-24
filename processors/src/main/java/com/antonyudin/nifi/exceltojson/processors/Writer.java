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


import java.math.BigDecimal;


public interface Writer {

	public enum Alignment {
		left, center, right
	}

	public record Style(Alignment alignment, Integer fontHeight) {
		public static Style aligned(final Alignment alignment, final Style style) {
			return new Style(alignment, (style != null? style.fontHeight(): null));
		}
		public static Style fontHeight(final Integer fontHeight, final Style style) {
			return new Style((style != null? style.alignment(): null), fontHeight);
		}
	}

	public void sheetStarted(final String name);
	public void sheetEnded();
	public void rowAdded();
	public void columnAdded(final String name, final boolean value, final Style style);
	public void columnAdded(final String name, final int value, final Style style);
	public void columnAdded(final String name, final String value, final Style style);
	public void columnAdded(final String name, final String value, final boolean formula, final Style style);
	public void columnAdded(final String name, final BigDecimal value, final Style style);
	public void mergeColumns(final int numberOfColumns);

	public enum Type {
		Boolean, String, Number
	}

}

