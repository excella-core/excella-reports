/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: OoPdfExporterTest.java 99 2010-01-19 06:45:54Z tomo-shibata $
 * $Revision: 99 $
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
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.OoPdfExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class OoPdfExporterTest extends ReportsWorkbookTest {

    public OoPdfExporterTest( String version) {
        super( version);
    }

    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    ConvertConfiguration configuration = null;

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {

        OoPdfExporter exporter = new OoPdfExporter();
        String filePath = null;

        Workbook wb = getWorkbook();

            wb = getWorkbook();
            configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
            filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
            exporter.setFilePath( filePath);
            try {
                exporter.output( wb, new BookData(), configuration);
                File file = new File( exporter.getFilePath());
                assertTrue( file.exists());
            } catch ( ExportException e) {
                e.printStackTrace();
                fail( e.toString());
            }

            // オプション指定
            wb = getWorkbook();
            configuration.addOption( "PermissionPassword", "pass");
            configuration.addOption( "RestrictPermissions", Boolean.TRUE);
            configuration.addOption( "Printing", 0);
            configuration.addOption( "Changes", 4);
            filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
            exporter.setFilePath( filePath);
            try {
                exporter.output( wb, new BookData(), configuration);
                File file = new File( exporter.getFilePath());
                assertTrue( file.exists());
            } catch ( ExportException e) {
                e.printStackTrace();
                fail( e.toString());
            }

            // 例外発生
            wb = getWorkbook();
            configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
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
                if ( e instanceof ExportException) {
                    // OK
                } else {
                    fail( e.toString());
                }
            }

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        OoPdfExporter exporter = new OoPdfExporter();
        assertEquals( "PDF", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        OoPdfExporter exporter = new OoPdfExporter();
        assertEquals( ".pdf", exporter.getExtention());
    }

}
