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
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.WorkbookTest;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.CellObject;
import org.bbreak.excella.reports.processor.ReportCreateHelper;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * {@link org.bbreak.excella.reports.tag.RemoveParamParser} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class RemoveParamParserTest extends ReportsWorkbookTest {

    /**
     * {@link org.bbreak.excella.reports.tag.RemoveParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)}
     * のためのテスト・メソッド。
     * 
     * @throws ParseException
     * @throws IOException
     */
    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public void testParseSheetCellObject( String version) throws ParseException, IOException {

        Workbook workbook = getWorkbook( version);
        Sheet sheet1 = workbook.getSheetAt( 0);

        RemoveParamParser parser = new RemoveParamParser();
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));

        // 解析処理
        List<ParsedReportInfo> results = parseSheet( parser, sheet1, reportsParserInfo);

        checkResult( new CellObject[] {new CellObject( 0, 0), new CellObject( 3, 2), new CellObject( 5, 3), new CellObject( 8, 1), new CellObject( 11, 4), new CellObject( 15, 2),
                new CellObject( 24, 3)}, results);

    }

    /**
     * {@link org.bbreak.excella.reports.tag.RemoveParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        RemoveParamParser parser = new RemoveParamParser();
        assertFalse( parser.useControlRow());
    }

    /**
     * {@link org.bbreak.excella.reports.tag.RemoveParamParser#RemoveParamPaser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testRemoveParamPaserString() {
        RemoveParamParser paser = new RemoveParamParser( "テスト");
        assertEquals( "テスト", paser.getTag());
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
