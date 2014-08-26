/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportsWorkbookTest.java 5 2009-06-22 07:55:44Z tomo-shibata $
 * $Revision: 5 $
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

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.WorkbookTest;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.junit.Assert;

public class ReportsWorkbookTest extends WorkbookTest {

    /** Excelファイルのバージョン */
    protected String version = null;

    public ReportsWorkbookTest( String version) {
        super( version);
        this.version = version;
    }

    protected Workbook getExpectedWorkbook() {

        Workbook workbook = null;

        String filename = this.getClass().getSimpleName();
        
        if ( version.equals( "2007")) {
            filename = filename + "_expected.xlsx";
        } else if ( version.equals( "2003")) {
            filename = filename + "_expected.xls";
        }
        
        
        URL url = this.getClass().getResource( filename);
        try {
            String path = URLDecoder.decode( url.getFile(), "UTF-8");
            
            workbook = getWorkbook(path);

        } catch ( UnsupportedEncodingException e) {
            Assert.fail();
        }
        return workbook;
    }

    
    protected List<ParsedReportInfo> parseSheet( ReportsTagParser<?> parser, Sheet sheet, ReportsParserInfo reportsParserInfo) throws ParseException {

        List<ParsedReportInfo> parsedList = new ArrayList<ParsedReportInfo>();
        
        for ( int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow( rowIndex);
            if ( row == null) {
                continue;
            }
            for ( int columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
                Cell cell = row.getCell( columnIndex);
                if ( cell == null) {
                    continue;
                }
                if ( parser.isParse( sheet, cell)) {
                    parsedList.add(parser.parse( sheet, cell, reportsParserInfo));
                }

            }

        }
        return parsedList;
    }
    
    
    protected Workbook getWorkbook(String filepath) {

        Workbook workbook = null;

        if ( filepath.endsWith( ".xlsx")) {
            try {
                workbook = new XSSFWorkbook( filepath);
            } catch ( IOException e) {
                Assert.fail();
            }
        } else if ( filepath.endsWith( ".xls")) {
            FileInputStream stream = null;
            try {
                stream = new FileInputStream( filepath);
            } catch ( FileNotFoundException e) {
                Assert.fail();
            }
            POIFSFileSystem fs = null;
            try {
                fs = new POIFSFileSystem( stream);
            } catch ( IOException e) {
                Assert.fail();
            }
            try {
                workbook = new HSSFWorkbook( fs);
            } catch ( IOException e) {
                Assert.fail();
            }
            try {
                stream.close();
            } catch ( IOException e) {
                Assert.fail();
            }
        }
         return workbook;
    }


}
