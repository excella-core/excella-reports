/*-
 * #%L
 * excella-reports
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.reports.processor;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.Map;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.BreakParamParser;
import org.bbreak.excella.reports.tag.ColRepeatParamParser;
import org.bbreak.excella.reports.tag.ImageParamParser;
import org.bbreak.excella.reports.tag.RemoveParamParser;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;
import org.bbreak.excella.reports.tag.SumParamParser;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.processor.ReportCreateHelper} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportCreateHelperTest {

    @Test
    public void testCreateDefaultParsers() {
        // デフォルト
        Map<String, ReportsTagParser<?>> parsers = ReportCreateHelper.createDefaultParsers();

        assertEquals( 9, parsers.size());
        for ( ReportsTagParser<?> parser : parsers.values()) {
            if ( parser.getTag().equals( SingleParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof SingleParamParser);
            } else if ( parser.getTag().equals( ImageParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof ImageParamParser);
            } else if ( parser.getTag().equals( RowRepeatParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof RowRepeatParamParser);
            } else if ( parser.getTag().equals( ColRepeatParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof ColRepeatParamParser);
            } else if ( parser.getTag().equals( BlockRowRepeatParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof BlockRowRepeatParamParser);
            } else if ( parser.getTag().equals( BlockColRepeatParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof BlockColRepeatParamParser);
            } else if ( parser.getTag().equals( SumParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof SumParamParser);
            } else if ( parser.getTag().equals( RemoveParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof RemoveParamParser);
            } else if ( parser.getTag().equals( BreakParamParser.DEFAULT_TAG)) {
                assertTrue( parser instanceof BreakParamParser);
            } else {
                fail( "デフォルトでないパーサー有");
            }
        }

    }

    @Test
    public void testCreateDefaultExporters() {
        // デフォルト
        Map<String, ReportBookExporter> exporters = ReportCreateHelper.createDefaultExporters();

        assertEquals( 1, exporters.size());
        for ( ReportBookExporter exporter : exporters.values()) {
            if ( exporter.getFormatType().equals( ExcelExporter.FORMAT_TYPE)) {
                assertTrue( exporter instanceof ExcelExporter);
            } else {
                fail( "デフォルトでないパーサー有");
            }
        }

    }
}
