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
