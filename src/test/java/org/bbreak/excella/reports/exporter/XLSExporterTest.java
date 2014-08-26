/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: XLSExporterTest.java 84 2009-10-30 04:07:19Z akira-yokoi $
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
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

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
 * {@link org.bbreak.excella.reports.exporter.XLSExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class XLSExporterTest {

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {
        XLSExporter exporter = new XLSExporter();

        ConvertConfiguration configuration = new ConvertConfiguration( exporter.getFormatType());

        String tmpDirPath = ReportsTestUtil.getTestOutputDir();

        String filePath = null;

        Workbook wb = new HSSFWorkbook();
        // XLS
        try {
            filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
            exporter.setFilePath( filePath);
            exporter.output( wb, new BookData(), configuration);
            File file = new File( exporter.getFilePath());
            assertTrue( file.exists());

        } catch ( ExportException e) {
            e.printStackTrace();
            fail( e.toString());
        }

        wb = new XSSFWorkbook();
        // XLSX
        try {
            filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
            exporter.setFilePath( filePath);
            exporter.output( wb, null, null);

            fail( "XLSXは解析不可");
        } catch ( IllegalArgumentException e) {
            // OK
        } catch ( ExportException e) {
            fail( e.toString());
        }

        // Exceptionを発生させる
        wb = new HSSFWorkbook();
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
     * {@link org.bbreak.excella.reports.exporter.XLSExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        XLSExporter exporter = new XLSExporter();
        assertEquals( "XLS", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.XLSExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        XLSExporter exporter = new XLSExporter();
        assertEquals( ".xls", exporter.getExtention());
    }

}
