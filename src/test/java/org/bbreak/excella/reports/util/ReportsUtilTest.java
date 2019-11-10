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

package org.bbreak.excella.reports.util;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportCreateHelper;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.ColRepeatParamParser;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.util.ReportsUtil} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class ReportsUtilTest extends ReportsWorkbookTest {

    public ReportsUtilTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getReportSheet(String, org.bbreak.excella.reports.model.ReportBook)} のためのテスト・メソッド。
     */
    @Test
    public void testGetReportSheet() {

        ReportBook reportBook = new ReportBook( "", "test", new ConvertConfiguration[] {new ConvertConfiguration( ExcelExporter.FORMAT_TYPE)});
        ReportSheet sheet1 = new ReportSheet( "Sheet1", "AAAA");
        reportBook.addReportSheet( sheet1);
        ReportSheet sheet2 = new ReportSheet( "Sheet1", "BBBB");
        reportBook.addReportSheet( sheet2);
        ReportSheet sheet3 = new ReportSheet( "Sheet1", "CCCC");
        reportBook.addReportSheet( sheet3);

        ReportSheet actual = ReportsUtil.getReportSheet( "AAAA", reportBook);
        assertEquals( sheet1, actual);

        actual = ReportsUtil.getReportSheet( "A", reportBook);
        assertNull( actual);

        actual = ReportsUtil.getReportSheet( "CCCC", reportBook);
        assertEquals( sheet3, actual);

    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getSheetNames(org.bbreak.excella.reports.model.ReportBook)} のためのテスト・メソッド。
     */
    @Test
    public void testGetSheetNames() {

        ReportBook reportBook = new ReportBook( "", "test", new ConvertConfiguration[] {new ConvertConfiguration( ExcelExporter.FORMAT_TYPE)});
        ReportSheet sheet1 = new ReportSheet( "Sheet1", "AAAA");
        reportBook.addReportSheet( sheet1);
        ReportSheet sheet2 = new ReportSheet( "Sheet1", "BBBB");
        reportBook.addReportSheet( sheet2);
        ReportSheet sheet3 = new ReportSheet( "Sheet1", "CCCC");
        reportBook.addReportSheet( sheet3);

        String[] expected = new String[] {"AAAA", "BBBB", "CCCC"};

        List<String> sheetNames = ReportsUtil.getSheetNames( reportBook);
        assertEquals( expected.length, sheetNames.size());

        assertArrayEquals( expected, sheetNames.toArray());

    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getParamValues(ParamInfo, String, List)} のためのテスト・メソッド。
     */
    @Test
    public void testGetParamValues() {

        // テストデータ
        ParamInfo info1 = createParamInfo( "", 0);

        ParamInfo info1$1 = createParamInfo( "", 1);
        ParamInfo info1$2 = createParamInfo( "", 2);
        ParamInfo info1$3 = createParamInfo( "", 3);
        ParamInfo info1$4 = createParamInfo( "", 4);
        ParamInfo info1$5 = createParamInfo( "", 5);
        ParamInfo info1$6 = createParamInfo( "", 6);

        Block block1 = new Block();
        block1.setParamboolean( true);
        block1.setParamBoolean( false);
        block1.setParambyte( ( byte) 1);
        block1.setParamByte( ( byte) 0);
        block1.setParamshort( ( short) 1);
        block1.setParamShort( ( short) 2);
        block1.setParamint( 100);
        block1.setParamIntger( 200);
        block1.setParamlong( 300L);
        block1.setParamLong( 400L);
        block1.setParamfloat( 0.1f);
        block1.setParamFloat( 0.2f);
        block1.setParamdouble( 0.3);
        block1.setParamDouble( 0.4);
        Calendar cal = Calendar.getInstance();
        cal.clear();
        cal.set( 2009, 6, 18);
        block1.setParamDate( cal.getTime());
        block1.setParamStr( "paramStr1");

        info1.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "p_br1", new ParamInfo[] {info1$1, info1$2, info1$3});
        info1.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "p_br2", new ParamInfo[] {info1$4, info1$5, info1$6});
        info1.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "p_br3", new Block[] {block1, block1, block1});

        ParamInfo info1$7 = createParamInfo( "", 7);
        ParamInfo info1$8 = createParamInfo( "", 8);
        ParamInfo info1$9 = createParamInfo( "", 9);
        ParamInfo info1$10 = createParamInfo( "", 10);

        info1.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "p_bc1", new ParamInfo[] {info1$7, info1$8});
        info1.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "p_bc2", new ParamInfo[] {info1$9, info1$10});
        info1.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "p_bc3", new Block[] {block1, block1});

        info1.addParam( "$CUSTOM", "custom1", Arrays.asList( new Object[] {"壱", 2, 3.0}));
        info1.addParam( "$CUSTOM", "custom2", Arrays.asList( new Object[] {"壱", 2, 3.0}));

        List<ReportsTagParser<?>> reportTagParserList = new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values());
        reportTagParserList.add( new CustomTagParser( "$CUSTOM"));

        List<Object> results = null;
        List<Object> expects = null;

        results = ReportsUtil.getParamValues( info1, "_p1", reportTagParserList);
        assertEquals( 1, results.size());
        expects = new ArrayList<Object>();
        expects.add( info1.getParam( SingleParamParser.DEFAULT_TAG, "_p1"));
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$R[]:_row2", reportTagParserList);
        assertEquals( 3, results.size());
        expects = new ArrayList<Object>();
        expects.add( 101);
        expects.add( 201);
        expects.add( 301);
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$C[]:_col3", reportTagParserList);
        assertEquals( 2, results.size());
        expects = new ArrayList<Object>();
        expects.add( 202);
        expects.add( 402);
        expects.addAll( Arrays.asList( new int[] {202, 402}));
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$BR[]:p_br1._p1", reportTagParserList);
        assertEquals( 3, results.size());
        expects = new ArrayList<Object>();
        expects.add( info1$1.getParam( SingleParamParser.DEFAULT_TAG, "_p1"));
        expects.add( info1$2.getParam( SingleParamParser.DEFAULT_TAG, "_p1"));
        expects.add( info1$3.getParam( SingleParamParser.DEFAULT_TAG, "_p1"));
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$BC[]:p_bc2._p3", reportTagParserList);
        assertEquals( 2, results.size());
        expects = new ArrayList<Object>();
        expects.add( info1$9.getParam( SingleParamParser.DEFAULT_TAG, "_p3"));
        expects.add( info1$10.getParam( SingleParamParser.DEFAULT_TAG, "_p3"));
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$CUSTOM:custom1", reportTagParserList);
        assertEquals( 3, results.size());
        expects = new ArrayList<Object>();
        expects.addAll( ( List<?>) info1.getParam( "$CUSTOM", "custom1"));
        assertTrue( expects.containsAll( results));

        results = ReportsUtil.getParamValues( info1, "$XX:XXX", reportTagParserList);
        assertEquals( 0, results.size());

        results = ReportsUtil.getParamValues( info1, "$BR[]:p_br3.paramint", reportTagParserList);
        assertEquals( 3, results.size());
        expects = new ArrayList<Object>();
        expects.add( block1.getParamint());
        expects.add( block1.getParamint());
        expects.add( block1.getParamint());
        assertTrue( expects.containsAll( results));

    }

    private ParamInfo createParamInfo( String param, int baseNum) {
        // テストデータ
        ParamInfo info = new ParamInfo();

        info.addParam( SingleParamParser.DEFAULT_TAG, param + "_p1", 10 + baseNum);
        info.addParam( SingleParamParser.DEFAULT_TAG, param + "_p2", 11 + baseNum);
        info.addParam( SingleParamParser.DEFAULT_TAG, param + "_p3", 12 + baseNum);

        info.addParam( RowRepeatParamParser.DEFAULT_TAG, param + "_row1", new int[] {100 + baseNum, 200 + baseNum, 300 + baseNum});
        info.addParam( RowRepeatParamParser.DEFAULT_TAG, param + "_row2", new int[] {101 + baseNum, 201 + baseNum, 301 + baseNum});
        info.addParam( RowRepeatParamParser.DEFAULT_TAG, param + "_row3", new int[] {102 + baseNum, 202 + baseNum, 302 + baseNum});

        info.addParam( ColRepeatParamParser.DEFAULT_TAG, param + "_col1", new int[] {200 + baseNum, 400 + baseNum});
        info.addParam( ColRepeatParamParser.DEFAULT_TAG, param + "_col2", new int[] {201 + baseNum, 401 + baseNum});
        info.addParam( ColRepeatParamParser.DEFAULT_TAG, param + "_col3", new int[] {202 + baseNum, 402 + baseNum});

        return info;

    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getSheetValues(org.bbreak.excella.reports.model.ReportBook, String)} のためのテスト・メソッド。
     */
    @Test
    public void testGetSheetValues() {

        ReportBook reportBook = new ReportBook( "", "test", new ConvertConfiguration[] {new ConvertConfiguration( ExcelExporter.FORMAT_TYPE)});
        ReportSheet sheet1 = new ReportSheet( "Sheet1", "AAAA");
        reportBook.addReportSheet( sheet1);
        sheet1.addParam( SingleParamParser.DEFAULT_TAG, "p1", "あいうえおAAAA");
        sheet1.addParam( SingleParamParser.DEFAULT_TAG, "p2", "かきくけこAAAA");
        sheet1.addParam( SingleParamParser.DEFAULT_TAG, "p3", "さしすせそAAAA");
        sheet1.addParam( RowRepeatParamParser.DEFAULT_TAG, "r1", new String[] {"A", "B", "C"});
        sheet1.addParam( RowRepeatParamParser.DEFAULT_TAG, "p1", new String[] {"X", "Y", "Z"});
        sheet1.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "br1", new ParamInfo[] {createParamInfo( "", 1), createParamInfo( "", 2)});

        ReportSheet sheet2 = new ReportSheet( "Sheet1", "BBBB");
        reportBook.addReportSheet( sheet2);
        sheet2.addParam( SingleParamParser.DEFAULT_TAG, "p1", "あいうえおBBBB");
        sheet2.addParam( SingleParamParser.DEFAULT_TAG, "p2", "かきくけこBBBB");
        // sheet2.addParam( SingleParamParser.DEFAULT_TAG, "p3","さしすせそBBBB");
        sheet2.addParam( RowRepeatParamParser.DEFAULT_TAG, "r1", new String[] {"A", "B", "C"});

        ReportSheet sheet3 = new ReportSheet( "Sheet1", "CCCC");
        reportBook.addReportSheet( sheet3);
        sheet3.addParam( SingleParamParser.DEFAULT_TAG, "p1", "あいうえおCCCC");
        // sheet3.addParam( SingleParamParser.DEFAULT_TAG, "p2","かきくけこCCCC");
        sheet3.addParam( SingleParamParser.DEFAULT_TAG, "p3", 3);

        List<Object> actual = null;

        List<ReportsTagParser<?>> parsers = new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values());

        actual = ReportsUtil.getSheetValues( reportBook, "p1", parsers);
        assertArrayEquals( new Object[] {"あいうえおAAAA", "あいうえおBBBB", "あいうえおCCCC"}, actual.toArray());

        actual = ReportsUtil.getSheetValues( reportBook, "p2", parsers);
        assertArrayEquals( new Object[] {"かきくけこAAAA", "かきくけこBBBB"}, actual.toArray());

        actual = ReportsUtil.getSheetValues( reportBook, "p3", parsers);
        assertArrayEquals( new Object[] {"さしすせそAAAA", 3}, actual.toArray());

        actual = ReportsUtil.getSheetValues( reportBook, "p4", parsers);
        assertEquals( 0, actual.size());

        actual = ReportsUtil.getSheetValues( reportBook, "$R[]:r1", parsers);
        assertEquals( 0, actual.size());

        actual = ReportsUtil.getSheetValues( reportBook, "XX", parsers);
        assertEquals( 0, actual.size());

        actual = ReportsUtil.getSheetValues( reportBook, "$BR[]:br1._p1", parsers);
        assertEquals( 0, actual.size());

    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getSumValue(ParamInfo, String, List)} のためのテスト・メソッド。
     */
    @Test
    public void testGetSumValue() {

        ParamInfo info = createParamInfo( "", 0);

        ParamInfo info$1 = new ParamInfo();
        info$1.addParam( SingleParamParser.DEFAULT_TAG, "p1", ( byte) 1);
        info$1.addParam( SingleParamParser.DEFAULT_TAG, "p2", ( short) 30);

        ParamInfo info$2 = new ParamInfo();
        info$2.addParam( SingleParamParser.DEFAULT_TAG, "p1", 100);
        info$2.addParam( SingleParamParser.DEFAULT_TAG, "p2", 300L);

        ParamInfo info$3 = new ParamInfo();
        info$3.addParam( SingleParamParser.DEFAULT_TAG, "p1", 3.5f);
        info$3.addParam( SingleParamParser.DEFAULT_TAG, "p2", 0.889);

        ParamInfo info$4 = new ParamInfo();
        info$4.addParam( SingleParamParser.DEFAULT_TAG, "p1", new BigInteger( "50000"));
        info$4.addParam( SingleParamParser.DEFAULT_TAG, "p2", new BigDecimal( "6.66"));

        info.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "br1", new ParamInfo[] {info$1, info$2, info$3, info$4});

        BigDecimal expected = (new BigDecimal( ( byte) 1 + 100).add( new BigDecimal( "3.5")).add( new BigDecimal( new BigInteger( "50000"))));
        BigDecimal decimal = ReportsUtil.getSumValue( info, "$BR[]:br1.p1", new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        assertEquals( expected.subtract( decimal).signum(), 0);
        expected = (new BigDecimal( ( short) 30 + 300L).add( BigDecimal.valueOf( 0.889)).add( new BigDecimal( "6.66")));
        decimal = ReportsUtil.getSumValue( info, "$BR[]:br1.p2", new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        assertEquals( expected, decimal);

        info = createParamInfo( "", 0);

        info$1 = new ParamInfo();
        info$1.addParam( SingleParamParser.DEFAULT_TAG, "p1", 100);
        info$1.addParam( SingleParamParser.DEFAULT_TAG, "p2", 200);

        info$2 = new ParamInfo();
        info$2.addParam( SingleParamParser.DEFAULT_TAG, "p1", 300);
        info$2.addParam( SingleParamParser.DEFAULT_TAG, "p2", 400);

        info$3 = new ParamInfo();
        info$3.addParam( SingleParamParser.DEFAULT_TAG, "p1", 500);
        info$3.addParam( SingleParamParser.DEFAULT_TAG, "p2", 600);

        info$4 = new ParamInfo();
        info$4.addParam( SingleParamParser.DEFAULT_TAG, "p1", 700);
        info$4.addParam( SingleParamParser.DEFAULT_TAG, "p2", 800);

        info.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "br1", new ParamInfo[] {info$1, info$2, info$3, info$4});

        expected = new BigDecimal( 100 + 300 + 500 + 700);
        decimal = ReportsUtil.getSumValue( info, "$BR[]:br1.p1", new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        assertEquals( expected.subtract( decimal).signum(), 0);
        expected = new BigDecimal( 200 + 400 + 600 + 800);
        decimal = ReportsUtil.getSumValue( info, "$BR[]:br1.p2", new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        assertEquals( expected, decimal);

        expected = new BigDecimal( 100 + 200 + 300);
        decimal = ReportsUtil.getSumValue( info, RowRepeatParamParser.DEFAULT_TAG + ":_row1", new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        assertEquals( expected, decimal);

    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getMergedAddress(Sheet, int, int)} のためのテスト・メソッド。
     */
    @Test
    public void testgetMergedAddress() {

        // ブック作成
        Workbook hssfWb = new HSSFWorkbook();

        // シート作成
        Sheet hssfSheet = hssfWb.createSheet( "testsheet");

        // 行作成
        Row hssfRow = hssfSheet.createRow( 0);

        // セル作成
        hssfRow.createCell( 0);
        hssfRow.createCell( 1);
        hssfRow.createCell( 2);

        // セル結合
        CellRangeAddress address1 = new CellRangeAddress( 0, 1, 1, 1);
        hssfSheet.addMergedRegion( address1);
        CellRangeAddress address2 = new CellRangeAddress( 0, 0, 2, 3);
        hssfSheet.addMergedRegion( address2);

        // テスト
        assertNull( ReportsUtil.getMergedAddress( hssfSheet, 0, 0));
        assertEquals( address1.toString(), ReportsUtil.getMergedAddress( hssfSheet, 0, 1).toString());
        assertEquals( address2.toString(), ReportsUtil.getMergedAddress( hssfSheet, 0, 2).toString());

        try {
            hssfWb.close();
        } catch(IOException e) {
        }
        
        // ブック作成
        Workbook xssfWb = new XSSFWorkbook();

        // シート作成
        Sheet xssfSheet = xssfWb.createSheet( "testsheet");

        // 行作成
        Row xssfRow = xssfSheet.createRow( 0);

        // セル作成
        xssfRow.createCell( 0);
        xssfRow.createCell( 1);
        xssfRow.createCell( 2);

        // セル結合
        address1 = new CellRangeAddress( 0, 1, 1, 1);
        xssfSheet.addMergedRegion( address1);
        address2 = new CellRangeAddress( 0, 0, 2, 3);
        xssfSheet.addMergedRegion( address2);

        // テスト
        assertNull( ReportsUtil.getMergedAddress( xssfSheet, 0, 0));
        assertEquals( address1.toString(), ReportsUtil.getMergedAddress( xssfSheet, 0, 1).toString());
        assertEquals( address2.toString(), ReportsUtil.getMergedAddress( xssfSheet, 0, 2).toString());

        try {
            xssfWb.close();
        } catch(IOException e) {
        }
    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsTagUtil#getCellIndex(java.lang.String, java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testGetCellIndex() {
        int[] pos = null;
        try {
            pos = ReportsUtil.getCellIndex( "3:5", "");
            assertEquals( 3, pos[0]);
            assertEquals( 5, pos[1]);

            try {
                pos = ReportsUtil.getCellIndex( "A:C", "");
                fail();
            } catch ( ParseException e) {
                assertTrue( true);
            }

            try {
                pos = ReportsUtil.getCellIndex( "TEST", "");
                fail();
            } catch ( ParseException e) {
                assertTrue( true);
            }

            try {
                pos = ReportsUtil.getCellIndex( "100", "");
                fail();
            } catch ( ParseException e) {
                assertTrue( true);
            }

        } catch ( ParseException e) {
            e.printStackTrace();
            fail();
        }
    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsTagUtil#getBlockCellValue(org.apache.poi.ss.usermodel.Sheet, int, int, int, int)} のためのテスト・メソッド。
     */
    @Test
    public void testGetBlockCellValue() {

        Workbook workbook = getWorkbook();
        Sheet sheet = workbook.getSheetAt( 0);

        Object[][] cellValues = ReportsUtil.getBlockCellValue( sheet, 0, 4, 0, 3);

        for ( int r = 0; r <= 4; r++) {
            for ( int c = 0; c <= 3; c++) {
                if ( cellValues[r][c] != null) {
                    try {
                        assertEquals( PoiUtil.getCellValue( sheet.getRow( r).getCell( c)), cellValues[r][c]);
                    } catch ( NullPointerException e) {
                        e.printStackTrace();
                        fail();
                    }
                } else {
                    assertTrue( sheet.getRow( r) == null || sheet.getRow( r).getCell( c) == null || sheet.getRow( r).getCell( c).getCellType() == CellType.BLANK);
                }

            }
        }
    }

    /**
     * {@link org.bbreak.excella.reports.util.ReportsTagUtil#getBlockCellStyle(org.apache.poi.ss.usermodel.Sheet, int, int, int, int)} のためのテスト・メソッド。
     */
    @Test
    public void testGetBlockCellStyle() {
        Workbook workbook = getWorkbook();
        Sheet sheet = workbook.getSheetAt( 0);

        Object[][] cellStyles = ReportsUtil.getBlockCellStyle( sheet, 0, 4, 0, 3);

        for ( int r = 0; r <= 4; r++) {
            for ( int c = 0; c <= 3; c++) {
                if ( cellStyles[r][c] != null) {
                    try {
                        assertTrue( sheet.getRow( r).getCell( c).getCellStyle().equals( cellStyles[r][c]));
                    } catch ( NullPointerException e) {
                        e.printStackTrace();
                        fail();
                    }
                } else {
                    assertTrue( sheet.getRow( r) == null || sheet.getRow( r).getCell( c) == null || sheet.getRow( r).getCell( c).getCellStyle() == null);
                }

            }
        }
    }
    
    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#getRowHeight(org.apache.poi.ss.usermodel.Sheet, int, int)} のためのテスト・メソッド。
     */
    @Test
    public void testGetRowHeight() {
        
        Workbook workbook = getWorkbook();
        Sheet sheet = workbook.getSheetAt( 0);
        
        Row row0 = sheet.getRow(0);
        float[] height = ReportsUtil.getRowHeight( sheet, 0, 19);

        // 値、書式の設定がある行の場合には
        // その行の高さが取得できること
        // ※値、書式の設定がある行を１行目(getRow(0))と定めてテスト
        assertTrue(height[0] > 0);

        assertEquals( height[0], row0.getHeightInPoints(), 0.01);
        
        // 値、書式の設定がない行の場合には
        // デフォルト値として「-1」を設定する
        // ※値、書式の設定がない行を２０行目(getRow(19))と定めてテスト
        assertEquals( height[19], -1, 0.01);
        
    }
    
    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#isEmptyRow( int[] rowCellTypes, Object[] rowCellValues, CellStyle[] rowCellStyles){}}のためのテスト・メソッド。
     */
    @Test
    public void testIsEmptyRow() {
      
        // テストシートの取得
        Workbook workbook = getWorkbook();
        Sheet sheet2 = workbook.getSheetAt( 1);

        // 開始列設定 ※初期値＝0（Ａ列）
        int defaultFromCellColIndex = 0;
        // 開始列設定 ※初期値＝0（１行目）
        int defaultFromCellRowIndex = 0;

        // 終了行、終了列を設定(終了行は４行目まであれば十分、終了列はバージョンに従う)
        int defaultToCellRowIndex = 5;
        int defaultToCellColIndex = SpreadsheetVersion.EXCEL97.getMaxColumns();
        if ( sheet2 instanceof XSSFSheet) {
            defaultToCellColIndex = SpreadsheetVersion.EXCEL2007.getMaxColumns();
        }

        // シート情報の保存[row][col]
        // タイプ、値、スタイル
        CellType[][] sheetCellTypes = ReportsUtil.getBlockCellType( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
        Object[][] sheetCellValues = ReportsUtil.getBlockCellValue( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
        CellStyle[][] sheetCellStyles = ReportsUtil.getBlockCellStyle( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);

        // テスト実行
        // １行目<rowIndex=0>
        // (タイプ＝設定なし、値＝設定あり、スタイル＝設定なし) → falseであること
        int rowIndex = 0;
        assertTrue( !ReportsUtil.isEmptyRow( sheetCellTypes[rowIndex], sheetCellValues[rowIndex], sheetCellStyles[rowIndex]));

        // ２行目<rowIndex=1>
        // (タイプ＝設定なし、値＝設定なし、スタイル＝設定あり) → falseであること
        rowIndex = 1;
        assertTrue( !ReportsUtil.isEmptyRow( sheetCellTypes[rowIndex], sheetCellValues[rowIndex], sheetCellStyles[rowIndex]));

        // ３行目<rowIndex=2>
        // (タイプ＝設定あり、値＝設定なし、スタイル＝設定なし) → falseであること
        rowIndex = 2;
        assertTrue( !ReportsUtil.isEmptyRow( sheetCellTypes[rowIndex], sheetCellValues[rowIndex], sheetCellStyles[rowIndex]));

        // ４行目<rowIndex=3>
        // (タイプ＝設定なし、値＝設定なし、スタイル＝設定なし) → trueであること
        rowIndex = 3;
        assertTrue( ReportsUtil.isEmptyRow( sheetCellTypes[rowIndex], sheetCellValues[rowIndex], sheetCellStyles[rowIndex]));
      
        
    }
    
    /**
     * {@link org.bbreak.excella.reports.util.ReportsUtil#isEmptyCell( int cellType, Object cellValue, CellStyle cellStyle}のためのテスト・メソッド。
     */
    @Test
    public void testIsEmptyCell() {
      
        // テストシートの取得
        Workbook workbook = getWorkbook();
        Sheet sheet2 = workbook.getSheetAt( 1);

        // 開始列設定 ※初期値＝0（Ａ列）
        int defaultFromCellColIndex = 0;
        // 開始列設定 ※初期値＝0（１行目）
        int defaultFromCellRowIndex = 0;

        // 終了行、終了列を設定(終了行は４行目まであれば十分、終了列はバージョンに従う)
        int defaultToCellRowIndex = 5;
        int defaultToCellColIndex = SpreadsheetVersion.EXCEL97.getMaxColumns();
        if ( sheet2 instanceof XSSFSheet) {
            defaultToCellColIndex = SpreadsheetVersion.EXCEL2007.getMaxColumns();
        }

        // シート情報の保存[row][col]
        // タイプ、値、スタイル
        CellType[][] sheetCellTypes = ReportsUtil.getBlockCellType( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
        Object[][] sheetCellValues = ReportsUtil.getBlockCellValue( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
        CellStyle[][] sheetCellStyles = ReportsUtil.getBlockCellStyle( sheet2, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);

        // テスト実行
        // １行目<rowIndex=0、colIndex=2>
        // (タイプ＝設定なし、値＝設定あり、スタイル＝設定なし) → falseであること
        int rowIndex = 0;
        int colIndex = 2;
        assertTrue( !ReportsUtil.isEmptyCell( sheetCellTypes[rowIndex][colIndex], sheetCellValues[rowIndex][colIndex], sheetCellStyles[rowIndex][colIndex]));

        // ２行目<rowIndex=1、colIndex=2>
        // (タイプ＝設定なし、値＝設定なし、スタイル＝設定あり) → falseであること
        rowIndex = 1;
        assertTrue( !ReportsUtil.isEmptyCell( sheetCellTypes[rowIndex][colIndex], sheetCellValues[rowIndex][colIndex], sheetCellStyles[rowIndex][colIndex]));

        // ３行目<rowIndex=2、colIndex=2>
        // (タイプ＝設定あり、値＝設定なし、スタイル＝設定なし) → falseであること
        rowIndex = 2;
        assertTrue( !ReportsUtil.isEmptyCell( sheetCellTypes[rowIndex][colIndex], sheetCellValues[rowIndex][colIndex], sheetCellStyles[rowIndex][colIndex]));

        // ４行目<rowIndex=3、colIndex=2>
        // (タイプ＝設定なし、値＝設定なし、スタイル＝設定なし) → trueであること
        rowIndex = 3;
        assertTrue( ReportsUtil.isEmptyCell( sheetCellTypes[rowIndex][colIndex], sheetCellValues[rowIndex][colIndex], sheetCellStyles[rowIndex][colIndex]));
        
    }
    

    class CustomTagParser extends ReportsTagParser<List<?>> {

        public CustomTagParser( String tag) {
            super( tag);
        }

        @Override
        public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
            return null;
        }

        @Override
        public boolean useControlRow() {
            return false;
        }

    }

    public class Block {
        private boolean paramboolean = false;

        private boolean paramBoolean = false;

        private byte parambyte = 0;

        private Byte paramByte = 0;

        private short paramshort = 0;

        private Short paramShort = 0;

        private int paramint = 0;

        private Integer paramIntger = 0;

        private long paramlong = 0L;

        private Long paramLong = 0L;

        private float paramfloat = 0.1f;

        private Float paramFloat = 0.1f;

        private double paramdouble = 0.1;

        private Double paramDouble = 0.1;

        private Date paramDate = null;

        private String paramStr = null;

        public boolean getParamboolean() {
            return paramboolean;
        }

        public void setParamboolean( boolean paramboolean) {
            this.paramboolean = paramboolean;
        }

        public Boolean getParamBoolean() {
            return paramBoolean;
        }

        public void setParamBoolean( Boolean paramBoolean) {
            this.paramBoolean = paramBoolean;
        }

        public byte getParambyte() {
            return parambyte;
        }

        public void setParambyte( byte parambyte) {
            this.parambyte = parambyte;
        }

        public Byte getParamByte() {
            return paramByte;
        }

        public void setParamByte( Byte paramByte) {
            this.paramByte = paramByte;
        }

        public short getParamshort() {
            return paramshort;
        }

        public void setParamshort( short paramshort) {
            this.paramshort = paramshort;
        }

        public Short getParamShort() {
            return paramShort;
        }

        public void setParamShort( Short paramShort) {
            this.paramShort = paramShort;
        }

        public int getParamint() {
            return paramint;
        }

        public void setParamint( int paramint) {
            this.paramint = paramint;
        }

        public Integer getParamIntger() {
            return paramIntger;
        }

        public void setParamIntger( Integer paramIntger) {
            this.paramIntger = paramIntger;
        }

        public long getParamlong() {
            return paramlong;
        }

        public void setParamlong( long paramlong) {
            this.paramlong = paramlong;
        }

        public Long getParamLong() {
            return paramLong;
        }

        public void setParamLong( Long paramLong) {
            this.paramLong = paramLong;
        }

        public float getParamfloat() {
            return paramfloat;
        }

        public void setParamfloat( float paramfloat) {
            this.paramfloat = paramfloat;
        }

        public Float getParamFloat() {
            return paramFloat;
        }

        public void setParamFloat( Float paramFloat) {
            this.paramFloat = paramFloat;
        }

        public double getParamdouble() {
            return paramdouble;
        }

        public void setParamdouble( double paramdouble) {
            this.paramdouble = paramdouble;
        }

        public Double getParamDouble() {
            return paramDouble;
        }

        public void setParamDouble( Double paramDouble) {
            this.paramDouble = paramDouble;
        }

        public Date getParamDate() {
            return paramDate;
        }

        public void setParamDate( Date paramDate) {
            this.paramDate = paramDate;
        }

        public String getParamStr() {
            return paramStr;
        }

        public void setParamStr( String paramStr) {
            this.paramStr = paramStr;
        }

    }
}
