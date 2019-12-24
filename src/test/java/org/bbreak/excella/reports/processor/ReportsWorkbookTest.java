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
