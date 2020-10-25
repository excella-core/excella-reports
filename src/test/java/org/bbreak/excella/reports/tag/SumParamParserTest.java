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

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.IOException;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.WorkbookTest;
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
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * {@link org.bbreak.excella.reports.tag.SumParamParser} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class SumParamParserTest extends ReportsWorkbookTest {

    /**
     * {@link org.bbreak.excella.reports.tag.SumParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)}
     * のためのテスト・メソッド。
     * 
     * @throws ParseException
     * @throws ReportsCheckException
     * @throws IOException
     */
    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public void testParseSheetCellObject( String version) throws ParseException, ReportsCheckException, IOException {

        // -----------------------
        // □[正常系]通常解析
        // -----------------------

        Workbook workbook = getWorkbook( version);

        Sheet sheet1 = workbook.getSheetAt( 0);

        ReportBook reportBook = new ReportBook("",  "test", new ConvertConfiguration[] {});
//        reportBook.setCopyTemplate( true);
        ReportSheet reportSheet = new ReportSheet( "Sheet1", "Sheet1");
        reportBook.addReportSheet( reportSheet);

        // テストデータ
        ParamInfo info = reportSheet.getParamInfo();

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

        SumParamParser parser = new SumParamParser();
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo.setReportBook( reportBook);
        reportsParserInfo.setParamInfo( reportSheet.getParamInfo());

        // 解析処理
        List<ParsedReportInfo> results = parseSheet( parser, sheet1, reportsParserInfo);

        // 処理結果のチェック
        checkResult( new CellObject[] {new CellObject( 4, 1), new CellObject( 7, 1)}, results);
        

        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }

        checkSheet( "Sheet1", sheet1, true, version);

    }

    /**
     * {@link org.bbreak.excella.reports.tag.SumParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        SumParamParser paser = new SumParamParser();
        assertFalse( paser.useControlRow());
    }

    /**
     * {@link org.bbreak.excella.reports.tag.SumParamParser#SumParamParser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testSumParamParserString() {
        SumParamParser paser = new SumParamParser( "テスト");
        assertEquals( "テスト", paser.getTag());
    }

    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel, String version)
            throws ReportsCheckException, IOException {

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook( version);
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, actualSheet, false);
        } finally {
            if ( outputExcel) {
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
    }

    private void checkResult( CellObject[] exceptedCells, List<ParsedReportInfo> results) {
        // 処理結果のチェック
        assertEquals( exceptedCells.length, results.size());
        CellObject[] actualCells = new CellObject[results.size()];

        for ( int i = 0; i < results.size(); i++) {
            ParsedReportInfo parsedReportInfo = results.get( i);
            assertEquals( exceptedCells[i].getRowIndex(), parsedReportInfo.getDefaultRowIndex());
            assertEquals( exceptedCells[i].getColIndex(), parsedReportInfo.getDefaultColumnIndex());
            actualCells[i] = new CellObject( parsedReportInfo.getRowIndex(), parsedReportInfo.getColumnIndex());
        }

        assertArrayEquals( exceptedCells, actualCells);

    }

}
