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

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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

        String filename = this.getClass().getSimpleName() + "_expected" + getSuffix();
        
        URL url = this.getClass().getResource( filename);
        try {
            String path = URLDecoder.decode( url.getFile(), "UTF-8");
            
            workbook = getWorkbook(path);

        } catch (IOException | EncryptedDocumentException e) {
            Assert.fail(e.toString());
        }
        return workbook;
    }

    protected String getSuffix() {
        if ( version.equals( "2007")) {
            return ".xlsx";
        }
        return ".xls";
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
    
    
    protected Workbook getWorkbook(String filepath) throws EncryptedDocumentException, IOException {
        return WorkbookFactory.create(new File(filepath));
    }


}
