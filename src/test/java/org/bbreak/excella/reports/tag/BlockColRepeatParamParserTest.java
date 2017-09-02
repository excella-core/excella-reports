/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: BlockColRepeatParamParserTest.java 173 2010-08-09 04:16:12Z ogiharasf $
 * $Revision: 173 $
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

import static org.junit.Assert.*;

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
 * {@link org.bbreak.excella.reports.tag.BlockColRepeatParamParser} のためのテスト・クラス。
 *
 */
public class BlockColRepeatParamParserTest extends ReportsWorkbookTest {
    
    

    /**
     * コンストラクタ
     * @param version エクセルバージョン
     */
    public BlockColRepeatParamParserTest( String version) {
        super(version);
    }

    
    @Test
    public void testParseSheetCellObject() throws ParseException {
        Workbook workbook = null;
        
        ReportBook reportBook = new ReportBook( "", "test", new ConvertConfiguration[] {});
        ReportSheet reportSheet1 = new ReportSheet( "sheet1", "Sheet1");
        reportBook.addReportSheet( reportSheet1);
        ReportSheet reportSheet2 = new ReportSheet( "sheet1", "Sheet2");
        reportBook.addReportSheet( reportSheet2);
        
        ReportSheet[] reportSheets = new ReportSheet[] {reportSheet1, reportSheet2};
        
        // -----------------------
        // テストデータ
        // -----------------------
        
        //子BCブロック1データ
        ParamInfo inBlockInfo1 = new ParamInfo();
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC1-C-A1", "inBC1-C-A2", "inBC1-C-A3"});
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBC1-C-B1", "inBC1-C-B1"});
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBC1-C-C1"});
        
        inBlockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC1-R-A1"});
        
        inBlockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBC1-S-DDD");
        
        //子BCブロック2データ
        ParamInfo inBlockInfo2 = new ParamInfo();
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC2-C-A1"});
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBC2-C-B1", "inBC2-C-B2"});
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBC2-C-C1", "inBC2-C-C2", "inBC2-C-C3"});
        
        inBlockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC2-R-A1", "inBC2-R-A2", "inBC2-R-A3"});
        
        inBlockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBC2-S-DDD");
        
        //子BCブロック3データ
        ParamInfo inBlockInfo3 = new ParamInfo();
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC3-C-A1", "inBC3-C-A2", "inBC3-C-A3"});
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBC3-C-B1", "inBC3-C-B2", "inBC3-C-B3", "inBC3-C-B4"});
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBC3-C-C1", "inBC3-C-C2", "inBC3-C-C3", "inBC3-C-C4", "inBC3-C-C5"});
        
        inBlockInfo3.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC3-R-A1", "inBC3-R-A2", "inBC3-R-A3", "inBC3-R-A4", "inBC3-R-A5"});
        
        inBlockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBC3-S-DDD");
        
        //子BCブロック4データ
        ParamInfo inBlockInfo4 = new ParamInfo();
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC4-C-A1", "inBC4-C-A2", "inBC4-C-A3", "inBC4-C-A4", "inBC4-C-A5"});
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBC4-C-B1", "inBC4-C-B2", "inBC4-C-B3", "inBC4-C-B4"});
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBC4-C-C1", "inBC4-C-C2", "inBC4-C-C3"});
        
        inBlockInfo4.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBC4-R-A1", "inBC4-R-A2"});
        
        inBlockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBC4-S-DDD");
        
        //BCブロック1データ
        ParamInfo blockInfo1 = new ParamInfo();
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BC1-C-A1", "BC1-C-A2", "BC1-C-A3", "BC1-C-A4", "BC1-C-A5"});
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BC1-C-B1", "BC1-C-B2", "BC1-C-B3", "BC1-C-B4"});
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BC1-C-C1", "BC1-C-C2", "BC1-C-C3"});
        
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BC1-R-A1", "BC1-R-A2", "BC1-R-A3", "BC1-R-A4", "BC1-R-A5"});
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BC1-R-B1", "BC1-R-B2", "BC1-R-B3", "BC1-R-B4"});
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BC1-R-C1", "BC1-R-C2", "BC1-R-C3"});
        
        blockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D", "BC1-S-DDD");
        blockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BC1-S-DDD2");
        
        blockInfo1.addParam(BlockColRepeatParamParser.DEFAULT_TAG, "inBC1", new Object[] {inBlockInfo1, inBlockInfo2});
        blockInfo1.addParam(BlockRowRepeatParamParser.DEFAULT_TAG, "inBC1", new Object[] {inBlockInfo1, inBlockInfo2});
        
        //BCブロック2データ
        ParamInfo blockInfo2 = new ParamInfo();
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BC2-C-A1", "BC2-C-A2", "BC2-C-A3", "BC2-C-A4"});
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BC2-C-B1", "BC2-C-B2", "BC2-C-B3", "BC2-C-B4", "BC2-C-B5"});
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BC2-C-C1", "BC2-C-C2", "BC2-C-C3"});
        
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BC2-R-A1", "BC2-R-A2", "BC2-R-A3", "BC2-R-A4"});
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BC2-R-B1", "BC2-R-B2", "BC2-R-B3", "BC2-R-B4", "BC2-R-B5"});
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BC2-R-C1", "BC2-R-C2", "BC2-R-C3"});
        
        blockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D", "BC2-S-DDD");
        blockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BC2-S-DDD2");
        
        blockInfo2.addParam(BlockColRepeatParamParser.DEFAULT_TAG, "inBC1", new Object[] {inBlockInfo3, inBlockInfo4});
        blockInfo2.addParam(BlockRowRepeatParamParser.DEFAULT_TAG, "inBC1", new Object[] {inBlockInfo3, inBlockInfo4});
        
        //BCブロック3データ
        ParamInfo blockInfo3 = new ParamInfo();
        blockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D", "BC2-S-DDD");
        blockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BC2-S-DDD2");

        //BCブロック4データ
        ParamInfo blockInfo4 = new ParamInfo();
        blockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D", "BC2-S-DDD");
        blockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BC2-S-DDD2");

        //BCブロック5データ
        ParamInfo blockInfo5 = new ParamInfo();
        blockInfo5.addParam(SingleParamParser.DEFAULT_TAG, "D", "BC5-S-DDD");
        blockInfo5.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BC5-S-DDD2");
        
        for (int i = 0; i < reportSheets.length; i++) {
            ParamInfo info = reportSheets[i].getParamInfo();
            if (i == 1) {
                info.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "BC1", new Object[]{blockInfo1, blockInfo2, blockInfo3, blockInfo4, blockInfo5});
            } else {
                info.addParam( BlockColRepeatParamParser.DEFAULT_TAG, "BC1", new Object[]{blockInfo1, blockInfo2});    
            }
        }

        BlockColRepeatParamParser parser = new BlockColRepeatParamParser();
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo.setReportBook( reportBook);
        reportsParserInfo.setParamInfo( reportSheets[0].getParamInfo());
        
        
        // 解析処理
        List<ParsedReportInfo> results = null;
        CellObject[] expectBeCells = null;
        CellObject[] expectAfCells = null;
        
        // -----------------------
        // □[正常系]オプション指定なし
        // BC-C
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet1 = workbook.getSheetAt( 0);
        results = parseSheet( parser, sheet1, reportsParserInfo);
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 3, 14)};
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
        // □[正常系]オプション指定なし
        // BC-C2
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet2 = workbook.getSheetAt( 1);
        
        results = parseSheet( parser, sheet2, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 3, 17)};
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
        // □[正常系]オプション指定なし
        // BC-R
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet3 = workbook.getSheetAt( 2);
        
        results = parseSheet( parser, sheet3, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 10, 6)};
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
        // □[正常系]オプション指定なし
        // BC-R2
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet4 = workbook.getSheetAt( 3);
        
        results = parseSheet( parser, sheet4, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 12, 6)};
        checkResult( expectBeCells, expectAfCells, results);
        
        checkSheet( "Sheet4", sheet4, true);
        
        
        // -----------------------
        // □[正常系]オプション指定なし
        // BC-CR
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet5 = workbook.getSheetAt( 4);
        
        results = parseSheet( parser, sheet5, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 4, 3)};
        expectAfCells = new CellObject[] {new CellObject( 8, 14)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "Sheet5", sheet5, true);
        
        
        // -----------------------
        // □[正常系]オプション指定なし
        // BC-BC
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet6 = workbook.getSheetAt( 5);
        
        results = parseSheet( parser, sheet6, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 8, 5), new CellObject( -1, -1), new CellObject( -1, -1)};
        expectAfCells = new CellObject[] {new CellObject( 12, 28), new CellObject( -1, -1), new CellObject( -1, -1)};
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkResult( expectBeCells, expectAfCells, results);
        
        checkSheet( "Sheet6", sheet6, true);
        
        
        // -----------------------
        // □[正常系]オプション指定なし
        // BC-BR
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet7 = workbook.getSheetAt( 6);
        
        results = parseSheet( parser, sheet7, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 8, 5)};
        expectAfCells = new CellObject[] {new CellObject( 17, 18)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet7", sheet7, true);
        
        
        // -----------------------
        // □[正常系]オプション指定
        // ・重複非表示
        // ・上限回数
        // -----------------------
        ReportsParserInfo reportsParserInfo8 = new ReportsParserInfo();
        reportsParserInfo8.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo8.setReportBook( reportBook);
        reportsParserInfo8.setParamInfo( reportSheets[1].getParamInfo());
        
        workbook = getWorkbook();
        Sheet sheet8 = workbook.getSheetAt( 7);
        
        results = parseSheet( parser, sheet8, reportsParserInfo8);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 3, 12)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet8", sheet8, true);
        
        
        // -----------------------
        // □[異常系]チェック
        // ・必須パラメータなし：fromCellなし、toCellなし
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet9 = workbook.getSheetAt( 8);
        try {
            results = parseSheet( parser, sheet9, reportsParserInfo);
            fail( "fromCell必須チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet9", sheet9, true);
        
        
        workbook = getWorkbook();
        Sheet sheet10 = workbook.getSheetAt( 9);
        try {
            results = parseSheet( parser, sheet10, reportsParserInfo);
            fail( "toCell必須チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet10", sheet10, true);
        
        
        // -----------------------
        // □[異常系]チェック
        // ・値不正：fromCell、toCell
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet11 = workbook.getSheetAt( 10);
        
        try {
            results = parseSheet( parser, sheet11, reportsParserInfo);
            fail( "fromCellの値の個数チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet11", sheet11, true);
        
        workbook = getWorkbook();
        Sheet sheet12 = workbook.getSheetAt( 11);
        
        try {
            results = parseSheet( parser, sheet12, reportsParserInfo);
            fail( "toCellの値の個数チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet12", sheet12, true);
        
        
        // -----------------------
        // □[異常系]チェック
        // ・値不正：fromCell、toCell
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet13 = workbook.getSheetAt( 12);
        
        try {
            results = parseSheet( parser, sheet13, reportsParserInfo);
            fail( "fromCellのマイナス値チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet13", sheet13, true);
        
        workbook = getWorkbook();
        Sheet sheet14 = workbook.getSheetAt( 13);
        
        try {
            results = parseSheet( parser, sheet14, reportsParserInfo);
            fail( "toCellのマイナス値チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet14", sheet14, true);
        
        
        // -----------------------
        // □[異常系]チェック
        // ・値不正：fromCell > toCell
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet15 = workbook.getSheetAt( 14);
        
        try {
            results = parseSheet( parser, sheet15, reportsParserInfo);
            fail( "fromCell > toCellチェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet15", sheet15, true);
        
        
        workbook = getWorkbook();
        Sheet sheet17 = workbook.getSheetAt( 16);
        
        try {
            results = parseSheet( parser, sheet17, reportsParserInfo);
            fail( "fromCell > toCellチェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet17", sheet17, true);
        
        
        // -----------------------
        // □[異常系]チェック
        // ・値不正：repeatNum
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet16 = workbook.getSheetAt( 15);
        
        try {
            results = parseSheet( parser, sheet16, reportsParserInfo);
            fail( "repeatNumの数値チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet16", sheet16, true);
        
        
        workbook = getWorkbook();
        Sheet sheet18 = workbook.getSheetAt( 17);
        
        try {
            results = parseSheet( parser, sheet18, reportsParserInfo);
            fail( "repeatNumの数値チェックにかかっていない");
        } catch (ParseException e) {
        }
        checkSheet( "sheet18", sheet18, true);
        
        // ----------------------------------
        // □[正常系]チェック
        //   １．指定範囲(toCell)に
        //       タイプ、値、スタイルのない行が存在
        //
        //   ２．指定範囲(toCell)に
        //       ある行において、値、スタイルの設定はないが、
        //       タイプの設定がある行が存在
        //
        //   ３．指定範囲外に列方向の結合セルが存在
        // ----------------------------------
        workbook = getWorkbook();
        Sheet sheet19 = workbook.getSheetAt( 18);
        
        results = parseSheet( parser, sheet19, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 6, 3)};
        expectAfCells = new CellObject[] {new CellObject( 6, 6)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet19", sheet19, true);
        
        // ----------------------------------
        // □[正常系]チェック
        //   行結合セルと列結合セル混在
        // ----------------------------------
        workbook = getWorkbook();
        Sheet sheet20 = workbook.getSheetAt( 19);
        
        results = parseSheet( parser, sheet20, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 11, 7)};
        expectAfCells = new CellObject[] {new CellObject( 19, 27)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet20", sheet20, true);
        
        // ----------------------------------
        // ■[異常系]チェック
        //   Rタグの同列後方に
        //   列方向の結合セルが存在する場合に
        //   発生する例外処理の確認
        // ----------------------------------
        workbook = getWorkbook();
        Sheet sheet21 = workbook.getSheetAt( 20);
        
        try {
            results = parseSheet( parser, sheet21, reportsParserInfo);
            fail( "例外処理が発生していない");
        } catch ( ParseException e) {
            // org.bbreak.excella.core.util.PoiUtil#getMergedAddress( Sheet sheet, CellRangeAddress rangeAddress)
            // でthrowした想定例外であることを確認する
            // ※補足
            //   BR、BCの場合には子タグから親タグに例外がスローされた時に
            //   ParseExceptionをさらに入れ子にして例外をスローするため、
            //   getCauseによるinstanceofは使用せずにgetMessageで評価を行いました。
            assertTrue( e.getMessage().contains("IllegalArgumentException"));
            assertTrue( e.getMessage().contains("There are crossing merged regions in the range."));
        }
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet21", sheet21, true);

        // ----------------------------------
        // ■[異常系]チェック
        //   Cタグの同行後方に
        //   行方向の結合セルが存在する場合に
        //   発生する例外処理の確認
        // ----------------------------------
        workbook = getWorkbook();
        Sheet sheet22 = workbook.getSheetAt( 21);
        
        try {
            results = parseSheet( parser, sheet22, reportsParserInfo);
            fail( "例外処理が発生していない");
        } catch ( ParseException e) {
            // org.bbreak.excella.core.util.PoiUtil#getMergedAddress( Sheet sheet, CellRangeAddress rangeAddress)
            // でthrowした想定例外であることを確認する
            // ※補足
            //   BR、BCの場合には子タグから親タグに例外がスローされた時に
            //   ParseExceptionをさらに入れ子にして例外をスローするため、
            //   getCauseによるinstanceofは使用せずにgetMessageで評価を行いました。
            assertTrue( e.getMessage().contains("IllegalArgumentException"));
            assertTrue( e.getMessage().contains("There are crossing merged regions in the range."));
        }
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet22", sheet22, true);

        // -----------------------
        // □[正常系]オプション指定
        // ・最低繰返回数(最低繰返数=paramInfo数)
        // -----------------------
        ReportsParserInfo reportsParserInfo23 = new ReportsParserInfo();
        reportsParserInfo23.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo23.setReportBook( reportBook);
        reportsParserInfo23.setParamInfo( reportSheets[0].getParamInfo());
        
        workbook = getWorkbook();
        Sheet sheet23 = workbook.getSheetAt( 22);
        
        results = parseSheet( parser, sheet23, reportsParserInfo23);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 3, 6)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet23", sheet23, true);
        
        // -----------------------
        // □[正常系]オプション指定
        // ・最低繰返回数(最低繰返数=paramInfo数+1)
        // -----------------------
        ReportsParserInfo reportsParserInfo24 = new ReportsParserInfo();
        reportsParserInfo24.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo24.setReportBook( reportBook);
        reportsParserInfo24.setParamInfo( reportSheets[0].getParamInfo());
        
        workbook = getWorkbook();
        Sheet sheet24 = workbook.getSheetAt( 23);
        
        results = parseSheet( parser, sheet24, reportsParserInfo24);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 3, 9)};
        checkResult( expectBeCells, expectAfCells, results);
        
        // 不要シートを削除
        if ( version.equals( "2007")) {
            int index = workbook.getSheetIndex( PoiUtil.TMP_SHEET_NAME);
            if ( index > 0) {
                workbook.removeSheetAt( index);
            }
        }
        
        checkSheet( "sheet24", sheet24, true);
    }
    
    
    /**
     * {@link org.bbreak.excella.reports.tag.BlockColRepeatParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        // -----------------------
        // □制御行使用有無：無
        // -----------------------
        BlockColRepeatParamParser parser = new BlockColRepeatParamParser();
        assertTrue( parser.useControlRow());
    }
    
    
    /**
     * {@link org.bbreak.excella.reports.tag.BlockColRepeatParamParser#BlockColRepeatParamParser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testBlockColRepeatParamParserString() {

        BlockColRepeatParamParser parser = new BlockColRepeatParamParser( "てすと");

        assertEquals( "てすと", parser.getTag());
    }
    
    
    
    
    

    
    
    
    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel) {

        // 期待値ブックの読み込み
        Workbook expectedWorkbook = getExpectedWorkbook();
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);
        
        expectedSheet.getPrintSetup().getCopies();

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
