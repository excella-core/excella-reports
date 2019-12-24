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

package org.bbreak.excella.reports.tag;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.CellObject;
import org.bbreak.excella.reports.processor.ReportCreateHelper;
import org.bbreak.excella.reports.processor.ReportsCheckException;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.tag.RowRepeatParamParser} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class RowRepeatParamParserTest extends ReportsWorkbookTest {

    public RowRepeatParamParserTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.tag.RowRepeatParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testParseSheetCellObject() {

        Workbook workbook = getWorkbook();
        Sheet sheet1 = workbook.getSheetAt( 0);

        // -----------------------
        // □[正常系]オプション指定なし
        // -----------------------

        ReportBook reportBook = new ReportBook("",  "test", new ConvertConfiguration[] {});
//        reportBook.setCopyTemplate( true);
        ReportSheet reportSheet1 = new ReportSheet( "sheet1", "Sheet1");
        reportBook.addReportSheet( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet( "sheet1", "Sheet2");
        reportBook.addReportSheet( reportSheet2);
        ReportSheet reportSheet3 = new ReportSheet( "sheet1", "Sheet3");
        reportBook.addReportSheet( reportSheet3);

        ReportSheet[] reportSheets = new ReportSheet[] {reportSheet1, reportSheet2, reportSheet3};

        for ( ReportSheet reportSheet : reportSheets) {
            ParamInfo info = reportSheet.getParamInfo();
            info.addParam( "$R[]", "A", new Object[] {"AA1", "AA1", "AA2", "AA2", "AA3"});
            info.addParam( "$R[]", "B", new Object[] {"BB1", "BB1", "BB2"});
            info.addParam( "$R[]", "C", new Object[] {"CC1", "CC2", "CC3", "CC4", "CC5"});
            info.addParam( "$", "D", "DDD");
        }

        RowRepeatParamParser parser = new RowRepeatParamParser();
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo.setReportBook( reportBook);
        reportsParserInfo.setParamInfo( reportSheets[0].getParamInfo());

        // 解析処理.
        List<ParsedReportInfo> results = null;
        try {
            results = parseSheet( parser, sheet1, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        CellObject[] expectBeCells = new CellObject[] {new CellObject( 0, 0), new CellObject( 2, 1)};
        CellObject[] expectAfCells = new CellObject[] {new CellObject( 4, 0), new CellObject( 4, 1)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }

        checkSheet( "Sheet1", sheet1, true);

        // -----------------------
        // □[正常系]オプション指定
        // ・行数改ページ
        // ・値変更改ページ
        // ・重複非表示
        // ・上限回数
        // ・セルシフト
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet2 = workbook.getSheetAt( 1);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet2, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        expectBeCells = new CellObject[] {new CellObject( 0, 0), new CellObject( 2, 1), new CellObject( 4, 2)};
        expectAfCells = new CellObject[] {new CellObject( 4, 0), new CellObject( 4, 1), new CellObject( 5, 2)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        checkSheet( "Sheet2", sheet2, true);

        // -----------------------
        // □[正常系]オプション指定
        // ・シートリンク
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet3 = workbook.getSheetAt( 2);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet3, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        expectBeCells = new CellObject[] {new CellObject( 0, 0), new CellObject( 0, 1), new CellObject( 0, 2), new CellObject( 17, 1), new CellObject( 18, 0)};
        expectAfCells = new CellObject[] {new CellObject( 2, 0), new CellObject( 1, 1), new CellObject( 4, 2), new CellObject( 19, 1), new CellObject( 20, 0)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        checkSheet( "Sheet3", sheet3, true);

        // -----------------------
        // ■[異常系]チェック
        // ・シートハイパーリンク設定有無と重複非表示は重複不可
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet4 = workbook.getSheetAt( 3);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet4, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        expectBeCells = new CellObject[] {new CellObject( 0, 0), new CellObject( 1, 0), new CellObject( 2, 0)};
        expectAfCells = new CellObject[] {new CellObject( 0, 0), new CellObject( 1, 0), new CellObject( 2, 0)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        checkSheet( "Sheet4", sheet4, true);

        Sheet sheet5 = workbook.getSheetAt( 4);
        // 解析処理
        try {
            parseSheet( parser, sheet5, reportsParserInfo);
            fail( "シートハイパーリンク設定有無と重複非表示は重複不可チェックにかかっていない");
        } catch ( ParseException e) {
        }

        // -----------------------
        // ■[異常系]チェック
        // ・エラーがあった場合
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet6 = workbook.getSheetAt( 5);
        // 解析処理
        try {
            parseSheet( parser, sheet6, reportsParserInfo);
            fail();
        } catch ( ParseException e) {
            assertTrue( e instanceof ParseException);
        }
        
        // ------------------------------------------------------------
        // ■[異常系]
        // ・同列後方に列方向の結合セルが存在する場合に
        //   PoiUtil.getMergedAddressメソッドにて
        //   定義した想定例外が発生することの確認を行う
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet7 = workbook.getSheetAt( 6);
        // 解析処理
        try {
            results = parseSheet( parser, sheet7, reportsParserInfo);
            fail( "想定例外が発生せず");
        } catch ( ParseException e) {
            // org.bbreak.excella.core.util.PoiUtil#getMergedAddress( Sheet sheet, CellRangeAddress rangeAddress)
            // でthrowした想定例外であることを確認する
            assertTrue( e.getCause() instanceof IllegalArgumentException);
            assertTrue( e.getMessage().contains("There are crossing merged regions in the range."));
        }
        
        // ------------------------------------------------------------
        // □[正常系]
        // ・行方向の結合セル（サイズ２）が存在する場合の正常終了確認
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet8 = workbook.getSheetAt( 7);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet8, reportsParserInfo);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail( e.toString());
        }
        
        expectBeCells = new CellObject[] {new CellObject(2,0), new CellObject(2,1), new CellObject(2,2), new CellObject(2,3), new CellObject(2,4), new CellObject(14,0)};
        expectAfCells = new CellObject[] {new CellObject(10,0), new CellObject(10,1), new CellObject(10,2), new CellObject(4,3), new CellObject(6,4), new CellObject(22,0)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        checkSheet( "Sheet8", sheet8, true);
        
        // ------------------------------------------------------------
        // □[正常系]
        // ・行方向の結合セル（サイズ３）が存在する場合の正常終了確認
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet9 = workbook.getSheetAt( 8);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet9, reportsParserInfo);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail( e.toString());
        }
        
        expectBeCells = new CellObject[] {new CellObject(3,0),new CellObject(3,1),new CellObject(3,2),new CellObject(3,3),new CellObject(3,4)};
        expectAfCells = new CellObject[] {new CellObject(15,0),new CellObject(15,1),new CellObject(15,2),new CellObject(6,3),new CellObject(9,4)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "Sheet9", sheet9, true);
        
        
        // ------------------------------------------------------------
        // ■[異常系]
        // ・タグセルが列行方向の結合セルである場合には
        //   PoiUtil.getMergedAddressメソッドにて
        //   定義した想定例外が発生することの確認を行う
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet10 = workbook.getSheetAt( 9);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet10, reportsParserInfo);
            fail( "想定例外が発生せず");
        } catch ( ParseException e) {
            // org.bbreak.excella.core.util.PoiUtil#getMergedAddress( Sheet sheet, CellRangeAddress rangeAddress)
            // でthrowした想定例外であることを確認する
            assertTrue( e.getCause() instanceof IllegalArgumentException);
            assertTrue( e.getMessage().contains("There are crossing merged regions in the range."));
        }

        // ------------------------------------------------------------
        // □[正常系]
        // ・行シフト先に結合セルがある
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet17 = workbook.getSheetAt( 16);
        results = null;
        try {
            results = parseSheet( parser, sheet17, reportsParserInfo);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail( e.toString());
        }
        checkSheet( "Sheet17", sheet17, true);
        
        // ------------------------------------------------------------
        // □[正常系]
        // ・最低繰返回数
        // ------------------------------------------------------------
        workbook = getWorkbook();
        Sheet sheet18 = workbook.getSheetAt( 17);
        // 解析処理
        results = null;
        try {
            results = parseSheet( parser, sheet18, reportsParserInfo);
        } catch ( ParseException e) {
            e.printStackTrace();
            fail( e.toString());
        }
        
        // 順にデータ数=最低繰返回数、データ数=最低繰返数-1,データ数=最低繰返数-1(結合セルが下にある),データ数=最低繰返数-1(結合セルを繰り返し)
        expectBeCells = new CellObject[] {new CellObject(1,1),new CellObject(2,2),new CellObject(2,4), new CellObject(4,3)};
        expectAfCells = new CellObject[] {new CellObject(5,1),new CellObject(5,2),new CellObject(5,4), new CellObject(10,3)};
        checkResult( expectBeCells, expectAfCells, results);

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "Sheet18", sheet18, true);
    }

    /**
     * {@link org.bbreak.excella.reports.tag.RowRepeatParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        // -----------------------
        // □制御行使用有無：無
        // -----------------------
        RowRepeatParamParser parser = new RowRepeatParamParser();
        assertFalse( parser.useControlRow());
    }

    /**
     * {@link org.bbreak.excella.reports.tag.RowRepeatParamParser#RowRepeatParamParser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRowRepeatParamParserString() {

        RowRepeatParamParser parser = new RowRepeatParamParser( "てすと");

        assertEquals( "てすと", parser.getTag());
    }

    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel) {

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, actualSheet, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            String tmpDirPath = ReportsTestUtil.getTestOutputDir();
            try {
                String filepath = null;
                Date now = new Date();
                if ( version.equals( "2007")) {
                    filepath = tmpDirPath + this.getClass().getSimpleName() + now.getTime() + ".xlsx";
                } else {
                    filepath = tmpDirPath + this.getClass().getSimpleName() + now.getTime() + ".xls";
                }
                PoiUtil.writeBook( actualSheet.getWorkbook(), filepath);

            } catch ( IOException e) {
                e.printStackTrace();
            }
        }
    }

    private void checkResult( CellObject[] exceptedBeforeCells, CellObject[] exceptedAfterCells, List<ParsedReportInfo> results) {
        // 処理結果のチェック
        assertEquals( exceptedAfterCells.length, results.size());

        CellObject[] actualBreforeCells = new CellObject[results.size()];
        CellObject[] actualAfterCells = new CellObject[results.size()];

        for ( int i = 0; i < results.size(); i++) {
            ParsedReportInfo parsedReportInfo = results.get( i);
            // 処理前値のチェック
            actualBreforeCells[i] = new CellObject( parsedReportInfo.getDefaultRowIndex(), parsedReportInfo.getDefaultColumnIndex());
            actualAfterCells[i] = new CellObject( parsedReportInfo.getRowIndex(), parsedReportInfo.getColumnIndex());
        }
        // 処理後値のチェック
        assertArrayEquals( exceptedBeforeCells, actualBreforeCells);
        assertArrayEquals( exceptedAfterCells, actualAfterCells);

    }
  
}
