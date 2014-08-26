/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportBookExporterTest.java 84 2009-10-30 04:07:19Z akira-yokoi $
 * $Revision: 84 $
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
package org.bbreak.excella.reports.exporter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.fail;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.WorkbookTest;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.ReportBookExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportBookExporterTest extends WorkbookTest {

    public ReportBookExporterTest( String version) {
        super( version);
        // TODO Auto-generated constructor stub
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ReportBookExporter#export(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData)} のためのテスト・メソッド。
     */
    @Test
    public void testExport() {

        Workbook book = getWorkbook();

        ReportBookExporter exporter = new ReportBookExporter() {

            @Override
            public String getExtention() {
                return null;
            }

            @Override
            public String getFormatType() {
                return null;
            }

            @Override
            public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
            }

        };

        exporter.setConfiguration( new ConvertConfiguration( ""));
        try {
            exporter.export( book, null);
        } catch ( ExportException e) {
            e.printStackTrace();
            fail( e.toString());
        }

        assertEquals( 4, book.getNumberOfSheets());

        for ( int i = 0; i < book.getNumberOfSheets(); i++) {
            String sheetName = book.getSheetName( i);
            assertFalse( sheetName.startsWith( PoiUtil.TMP_SHEET_NAME));
        }

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ReportBookExporter#setup()} のためのテスト・メソッド。
     */
    @Test
    public void testSetUp() {
        ReportBookExporter exporter = new ReportBookExporter() {

            @Override
            public String getExtention() {
                return null;
            }

            @Override
            public String getFormatType() {
                return null;
            }

            @Override
            public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
            }

        };
        exporter.setup();

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ReportBookExporter#tearDown()} のためのテスト・メソッド。
     */
    @Test
    public void testTearDown() {
        ReportBookExporter exporter = new ReportBookExporter() {

            @Override
            public String getExtention() {
                return null;
            }

            @Override
            public String getFormatType() {
                return null;
            }

            @Override
            public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
            }

        };
        exporter.tearDown();

    }

}
