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

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.XLSXExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class XLSXExporterTest {

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSXExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {
        XLSXExporter exporter = new XLSXExporter();

        ConvertConfiguration configuration = new ConvertConfiguration( exporter.getFormatType());

        String tmpDirPath = ReportsTestUtil.getTestOutputDir();

        String filePath = null;

        Workbook wb = new HSSFWorkbook();
        // XLS
        try {
            filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
            exporter.setFilePath( filePath);
            exporter.output( wb, new BookData(), configuration);
            fail( "XLSは解析不可");

        } catch ( IllegalArgumentException e) {
            // OK
        } catch ( ExportException e) {
            e.printStackTrace();
            fail( e.toString());
        }

        wb = new XSSFWorkbook();
        // XLSX
        try {
            filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
            exporter.setFilePath( filePath);
            exporter.output( wb, new BookData(), configuration);
            File file = new File( exporter.getFilePath());
            assertTrue( file.exists());

        } catch ( IllegalArgumentException e) {
            // OK
        } catch ( ExportException e) {
            e.printStackTrace();
            fail( e.toString());
        }

        // Exceptionを発生させる
        wb = new XSSFWorkbook();
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        try {
            exporter.output( wb, new BookData(), configuration);
        } catch ( ExportException e) {
            fail( e.toString());
        }
        File file = new File( exporter.getFilePath());
        file.setReadOnly();
        try {
            exporter.output( wb, new BookData(), configuration);
            fail( "例外未発生");
        } catch ( Exception e) {
            if ( e.getCause() instanceof IOException) {
                // OK
            } else {
                fail( e.toString());
            }
        }

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSXExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        XLSExporter exporter = new XLSExporter();
        assertEquals( "XLS", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSXExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        XLSExporter exporter = new XLSExporter();
        assertEquals( ".xls", exporter.getExtention());
    }

}
