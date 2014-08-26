/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ExcelExporterTest.java 12 2009-06-24 02:05:33Z tomo-shibata $
 * $Revision: 12 $
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

import java.io.File;
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
 * {@link org.bbreak.excella.reports.exporter.ExcelExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 *
 */
public class ExcelExporterTest {
    
    private Workbook hssfWb = new HSSFWorkbook();

    private Workbook xssfWb = new XSSFWorkbook();
    
    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();
    
    ConvertConfiguration configuration = null;
    
    
    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)} のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {
        
        ExcelExporter exporter = new ExcelExporter();
        String filePath = null;
        
        //XLS
        configuration = new ConvertConfiguration(ExcelExporter.EXTENTION_XLS);
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        try {
            exporter.output( hssfWb, new BookData(), configuration);
            File file = new File(exporter.getFilePath());
            assertTrue( file.exists());
        } catch ( ExportException e) {
            e.printStackTrace();
        }
        
        
        //XLSX
        configuration = new ConvertConfiguration(ExcelExporter.EXTENTION_XLSX);
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);
        try {
            exporter.output( xssfWb, new BookData(), configuration);
            File file = new File(exporter.getFilePath());
            assertTrue( file.exists());
            
        } catch ( ExportException e) {
            e.printStackTrace();
        }
        
        
        
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        ExcelExporter exporter = new ExcelExporter();
        assertEquals( "EXCEL", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.ExcelExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        ExcelExporter exporter = new ExcelExporter();
        assertEquals( "", exporter.getExtention());
    }
    
}
