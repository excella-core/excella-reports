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

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.fail;

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
import org.junit.jupiter.api.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.XLSXExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class XLSXExporterTest {

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSXExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     * 
     * @throws ExportException
     */
    @Test
    public void testOutput() throws ExportException {
        XLSXExporter exporter = new XLSXExporter();

        ConvertConfiguration configuration = new ConvertConfiguration( exporter.getFormatType());

        String tmpDirPath = ReportsTestUtil.getTestOutputDir();

        String filePath = null;

        Workbook xlsWorkbook = new HSSFWorkbook();
        // XLS
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        assertThrows( IllegalArgumentException.class, () -> exporter.output( xlsWorkbook, new BookData(), configuration));

        Workbook xlsxWorkbook = new XSSFWorkbook();
        // XLSX
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        assertDoesNotThrow( () -> exporter.output( xlsxWorkbook, new BookData(), configuration));
        File file = new File( exporter.getFilePath());
        assertTrue( file.exists());

        // Exceptionを発生させる
        Workbook xlsxWorkbook2 = new XSSFWorkbook();
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        exporter.output( xlsxWorkbook2, new BookData(), configuration);
        file = new File( exporter.getFilePath());
        file.setReadOnly();
        ExportException ee = assertThrows( ExportException.class, () -> exporter.output( xlsxWorkbook2, new BookData(), configuration));
        if ( !(ee.getCause() instanceof IOException)) {
            fail(ee);
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
