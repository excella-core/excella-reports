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

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.bbreak.excella.reports.tag.SingleParamParser;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.processor.ReportsParserInfo} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportsParserInfoTest {

    /**
     * @throws java.lang.Exception
     */
    @BeforeClass
    public static void setUpBeforeClass() throws Exception {
    }

    /**
     * @throws java.lang.Exception
     */
    @Before
    public void setUp() throws Exception {
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#getParamInfo()} のためのテスト・メソッド。
     */
    @Test
    public void testGetParamInfo() {
        ReportsParserInfo info = new ReportsParserInfo();
        ParamInfo paramInfo = new ParamInfo();
        
        info.setParamInfo(paramInfo);
        
        assertEquals( paramInfo, info.getParamInfo());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#setParamInfo(org.bbreak.excella.reports.model.ParamInfo)} のためのテスト・メソッド。
     */
    @Test
    public void testSetParamInfo() {
        ReportsParserInfo info = new ReportsParserInfo();
        ParamInfo paramInfo = new ParamInfo();
        
        info.setParamInfo(paramInfo);
        
        assertEquals( paramInfo, info.getParamInfo());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#getReportParsers()} のためのテスト・メソッド。
     */
    @Test
    public void testGetReportParsers() {
        ReportsParserInfo info = new ReportsParserInfo();
        
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>(ReportCreateHelper.createDefaultParsers().values());
        info.setReportParsers( reportParsers);
        
        assertArrayEquals( reportParsers.toArray(), info.getReportParsers().toArray());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#setReportParsers(java.util.List)} のためのテスト・メソッド。
     */
    @Test
    public void testSetReportParsers() {
        ReportsParserInfo info = new ReportsParserInfo();
        
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>(ReportCreateHelper.createDefaultParsers().values());
        info.setReportParsers( reportParsers);
        
        assertArrayEquals( reportParsers.toArray(), info.getReportParsers().toArray());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#createChildParserInfo(org.bbreak.excella.reports.model.ParamInfo)} のためのテスト・メソッド。
     */
    @Test
    public void testCreateChildParserInfo() {
        
        ReportsParserInfo info = new ReportsParserInfo();
        ReportBook reportBook = new ReportBook("", "", new ConvertConfiguration[]{});
        info.setReportBook( reportBook);
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>(ReportCreateHelper.createDefaultParsers().values());
        info.setReportParsers( reportParsers);
        
        ParamInfo paramInfo = new ParamInfo();
        info.setParamInfo(paramInfo);
        
        ParamInfo childParamInfo = new ParamInfo();
        
        ReportsParserInfo parserInfo = info.createChildParserInfo( childParamInfo);
        
        assertEquals( reportBook, parserInfo.getReportBook());
        assertArrayEquals( reportParsers.toArray(),  parserInfo.getReportParsers().toArray());
        assertEquals(childParamInfo, parserInfo.getParamInfo());
        
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#getMatchTagParser(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell)} のためのテスト・メソッド。
     */
    @Test
    public void testGetMatchTagParser() {
        
        ReportsParserInfo info = new ReportsParserInfo();
        ReportBook reportBook = new ReportBook("", "", new ConvertConfiguration[]{});
        info.setReportBook( reportBook);
        List<ReportsTagParser<?>> reportParsers = new ArrayList<ReportsTagParser<?>>(ReportCreateHelper.createDefaultParsers().values());
        info.setReportParsers( reportParsers);
        
        Workbook hssfWb = new HSSFWorkbook();
        Workbook xssfWb = new XSSFWorkbook();
        Sheet hssfSheet = hssfWb.createSheet("new sheet");
        Sheet xssfSheet = xssfWb.createSheet("new sheet");
        Cell hssfCell0 = hssfSheet.createRow( 0).createCell( 0);
        Cell hssfCell1 = hssfSheet.createRow( 1).createCell( 0);
        hssfCell0.setCellValue( "${HSSF}");
        hssfCell1.setCellValue( "$TEST{HSSF}");
        Cell xssfCell0 = xssfSheet.createRow( 0).createCell( 0);
        Cell xssfCell1 = xssfSheet.createRow( 1).createCell( 0);
        xssfCell0.setCellValue( "${XSSF}");
        xssfCell1.setCellValue( "$TEST{XSSF}");
        

        TagParser<?> parser = null;
        try {
            parser = info.getMatchTagParser( hssfSheet, hssfCell0);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail(e.toString());
        }
        assertEquals( SingleParamParser.class, parser.getClass());


        try {
            parser = info.getMatchTagParser( hssfSheet, hssfCell1);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail(e.toString());
        }
        assertNull( parser);
        
        
        try {
            parser = info.getMatchTagParser( xssfSheet, xssfCell0);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail(e.toString());
        }
        assertEquals( SingleParamParser.class, parser.getClass());
        
        try {
            parser = info.getMatchTagParser( xssfSheet, xssfCell1);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail(e.toString());
        }
        assertNull( parser);

        try {
            hssfWb.close();
        } catch(IOException e) {
        }
        try {
            xssfWb.close();
        } catch(IOException e) {
        }
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#getReportBook()} のためのテスト・メソッド。
     */
    @Test
    public void testGetReportBook() {
        ReportsParserInfo info = new ReportsParserInfo();
        
        ReportBook reportBook = new ReportBook("", "", new ConvertConfiguration[]{});
        info.setReportBook( reportBook);
        
        assertEquals( reportBook, info.getReportBook());
    }

    /**
     * {@link org.bbreak.excella.reports.processor.ReportsParserInfo#setReportBook(org.bbreak.excella.reports.model.ReportBook)} のためのテスト・メソッド。
     */
    @Test
    public void testSetReportBook() {
        ReportsParserInfo info = new ReportsParserInfo();
        
        ReportBook reportBook = new ReportBook("", "", new ConvertConfiguration[]{});
        info.setReportBook( reportBook);
        
        assertEquals( reportBook, info.getReportBook());
    }

}
