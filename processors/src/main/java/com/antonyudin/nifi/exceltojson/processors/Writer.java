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

	public enum ImageType {
		png, jpeg
	}

	public enum Alignment {
		left, center, right
	}

	public record Style(
		Alignment alignment,
		Integer fontHeight,
		Boolean fontWeightBold,
		String color,
		String fillColor,
		Integer width
	) {

		public static Style of() {
			return new Style(null, null, null, null, null, null);
		}

		public static class Builder {

			private Alignment alignment;
			private Integer fontHeight;
			private Boolean fontWeightBold;
			private String color;
			private String fillColor;
			private Integer width;

			public Builder aligned(final Alignment alignment) {
				this.alignment = alignment;
				return this;
			}

			public Builder fontHeight(final Integer fontHeight) {
				this.fontHeight = fontHeight;
				return this;
			}

			public Builder fontWeightBold(final boolean bold) {
				this.fontWeightBold = bold;
				return this;
			}

			public Builder color(final String color) {
				this.color = color;
				return this;
			}

			public Builder fillColor(final String fillColor) {
				this.fillColor = fillColor;
				return this;
			}

			public Builder width(final Integer width) {
				this.width = width;
				return this;
			}

			public Style build() {
				return new Style(
					alignment,
					fontHeight,
					fontWeightBold,
					color,
					fillColor,
					width
				);
			}

			public boolean isEmpty() {
				if (alignment != null)
					return false;
				if (fontHeight != null)
					return false;
				if (fontWeightBold != null)
					return false;
				if (color != null)
					return false;
				if (fillColor != null)
					return false;
				if (width != null)
					return false;
				return true;
			}
		}
	}


	public void startSheet(final String name, final Style style, final boolean selected, final boolean active);
	public void endSheet() throws Exception;
	public void addRow();
	public void addColumn(final String name, final boolean value, final Style style);
	public void addColumn(final String name, final int value, final Style style);
	public void addColumn(final String name, final String value, final Style style);
	public void addColumn(final String name, final String value, final boolean formula, final Style style);
	public void addColumn(final String name, final BigDecimal value, final Style style);
	public void mergeColumns(final int numberOfColumns);
	public void setPrintArea(final String reference);
	public void addImage(final String reference, final String image, final ImageType type, final double scale) throws Exception;

}

