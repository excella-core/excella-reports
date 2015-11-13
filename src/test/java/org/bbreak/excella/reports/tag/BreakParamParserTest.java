/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id$
 * $Revision$
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
 * version 3 along with ExCella Reports .  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/

package org.bbreak.excella.reports.tag;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.fail;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.CellObject;
import org.bbreak.excella.reports.processor.ReportCreateHelper;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.tag.BreakParamParser} のためのテスト・クラス。
 * 
 * @author T.Maruyama
 */
public class BreakParamParserTest extends ReportsWorkbookTest {

    /**
     * コンストラクタ
     * 
     * @param version Excelバージョン
     */
    public BreakParamParserTest( String version) {
        super( version);
    }

    /**
     * {@link org.bbreak.excella.reports.tag.BreakParamParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)} のためのテスト・メソッド。
     */
    @Test
    public void testParseSheetCellObject() {

        Workbook workbook = getWorkbook();
        Sheet sheet1 = workbook.getSheetAt( 0);

        BreakParamParser parser = new BreakParamParser();
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));

        List<ParsedReportInfo> results = null;
        // 解析処理
        try {
            results = parseSheet( parser, sheet1, reportsParserInfo);
        } catch ( ParseException e) {
            fail( e.toString());
        }

        checkResult(
            new CellObject[] {new CellObject( 0, 0), new CellObject( 3, 2), new CellObject( 5, 3), new CellObject( 8, 1), new CellObject( 11, 4), new CellObject( 15, 2), new CellObject( 24, 3)},
            results);

    }

    /**
     * {@link org.bbreak.excella.reports.tag.BreakParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        BreakParamParser parser = new BreakParamParser();
        assertFalse( parser.useControlRow());
    }

    /**
     * {@link org.bbreak.excella.reports.tag.BreakParamParser#BreakParamPaser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testBreakParamPaserString() {
        BreakParamParser paser = new BreakParamParser( "テスト");
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
