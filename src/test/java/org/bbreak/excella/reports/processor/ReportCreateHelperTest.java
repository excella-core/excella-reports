/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportCreateHelperTest.java 55 2009-10-19 11:40:05Z a-hoshino $
 * $Revision: 55 $
 *
 * This file is part of ExCella Reports.
 *
 * ExCella Reports is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * ExCella Reports is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the COPYING.LESSER file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with ExCella Reports.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.reports.processor;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.util.Map;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
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

        assertEquals( 8, parsers.size());
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
            } else {
                fail( "デフォルトでないパーサー有");
            }
        }

    }

    @Test
    public void testCreateDefaultExporters() {
        // デフォルト
        Map<String, ReportBookExporter> exporters = ReportCreateHelper.createDefaultExporters();

        assertEquals( 2, exporters.size());
        for ( ReportBookExporter exporter : exporters.values()) {
            if ( exporter.getFormatType().equals( ExcelExporter.FORMAT_TYPE)) {
                assertTrue( exporter instanceof ExcelExporter);
            } else {
                fail( "デフォルトでないパーサー有");
            }
        }

    }
}
