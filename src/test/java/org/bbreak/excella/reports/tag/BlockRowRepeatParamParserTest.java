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
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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


public class BlockRowRepeatParamParserTest  extends ReportsWorkbookTest {
    
    /**
     * コンストラクタ
     * @param version エクセルバージョン
     */
    public BlockRowRepeatParamParserTest( String version) {
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
        
        //子ブロック1データ
        ParamInfo inBlockInfo1 = new ParamInfo();
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR1-C-A1", "inBR1-C-A2", "inBR1-C-A3"});
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBR1-C-B1", "inBR1-C-B1"});
        inBlockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBR1-C-C1"});
        
        inBlockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR1-R-A1"});
        
        inBlockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBR1-S-DDD");
        
        //子ブロック2データ
        ParamInfo inBlockInfo2 = new ParamInfo();
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR2-C-A1"});
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBR2-C-B1", "inBR2-C-B2"});
        inBlockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBR2-C-C1", "inBR2-C-C2", "inBR2-C-C3"});
        
        inBlockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR2-R-A1", "inBR2-R-A2", "inBR2-R-A3"});
        
        inBlockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBR2-S-DDD");
        
        //子ブロック3データ
        ParamInfo inBlockInfo3 = new ParamInfo();
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR3-C-A1", "inBR3-C-A2", "inBR3-C-A3"});
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBR3-C-B1", "inBR3-C-B2", "inBR3-C-B3", "inBR3-C-B4"});
        inBlockInfo3.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBR3-C-C1", "inBR3-C-C2", "inBR3-C-C3", "inBR3-C-C4", "inBR3-C-C5"});
        
        inBlockInfo3.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR3-R-A1", "inBR3-R-A2", "inBR3-R-A3", "inBR3-R-A4", "inBR3-R-A5"});
        
        inBlockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBR3-S-DDD");
        
        //子ブロック4データ
        ParamInfo inBlockInfo4 = new ParamInfo();
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR4-C-A1", "inBR4-C-A2", "inBR4-C-A3", "inBR4-C-A4", "inBR4-C-A5"});
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"inBR4-C-B1", "inBR4-C-B2", "inBR4-C-B3", "inBR4-C-B4"});
        inBlockInfo4.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"inBR4-C-C1", "inBR4-C-C2", "inBR4-C-C3"});
        
        inBlockInfo4.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"inBR4-R-A1", "inBR4-R-A2"});
        
        inBlockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D", "inBR4-S-DDD");
        
        //BRブロック1データ
        ParamInfo blockInfo1 = new ParamInfo();
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BR1-C-A1", "BR1-C-A2", "BR1-C-A3", "BR1-C-A4", "BR1-C-A5"});
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BR1-C-B1", "BR1-C-B2", "BR1-C-B3", "BR1-C-B4"});
        blockInfo1.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BR1-C-C1", "BR1-C-C2", "BR1-C-C3"});
        
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BR1-R-A1", "BR1-R-A2", "BR1-R-A3", "BR1-R-A4", "BR1-R-A5"});
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BR1-R-B1", "BR1-R-B2", "BR1-R-B3", "BR1-R-B4"});
        blockInfo1.addParam( RowRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BR1-R-C1", "BR1-R-C2", "BR1-R-C3"});
        
        blockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D", "BR1-S-DDD");
        blockInfo1.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BR1-S-DDD2");
        
        blockInfo1.addParam(BlockColRepeatParamParser.DEFAULT_TAG, "inBR1", new Object[] {inBlockInfo1, inBlockInfo2});
        blockInfo1.addParam(BlockRowRepeatParamParser.DEFAULT_TAG, "inBR1", new Object[] {inBlockInfo1, inBlockInfo2});
        
        //BCブロック2データ
        ParamInfo blockInfo2 = new ParamInfo();
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BR2-C-A1", "BR2-C-A2", "BR2-C-A3", "BR2-C-A4"});
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BR2-C-B1", "BR2-C-B2", "BR2-C-B3", "BR2-C-B4", "BR2-C-B5"});
        blockInfo2.addParam( ColRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BR2-C-C1", "BR2-C-C2", "BR2-C-C3"});
        
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "A", new Object[] {"BR2-R-A1", "BR2-R-A2", "BR2-R-A3", "BR2-R-A4"});
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "B", new Object[] {"BR2-R-B1", "BR2-R-B2", "BR2-R-B3", "BR2-R-B4", "BR2-R-B5"});
        blockInfo2.addParam( RowRepeatParamParser.DEFAULT_TAG, "C", new Object[] {"BR2-R-C1", "BR2-R-C2", "BR2-R-C3"});
        
        blockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D", "BR2-S-DDD");
        blockInfo2.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BR2-S-DDD2");
        
        blockInfo2.addParam(BlockColRepeatParamParser.DEFAULT_TAG, "inBR1", new Object[] {inBlockInfo3, inBlockInfo4});
        blockInfo2.addParam(BlockRowRepeatParamParser.DEFAULT_TAG, "inBR1", new Object[] {inBlockInfo3, inBlockInfo4});
        
        //BCブロック3データ
        ParamInfo blockInfo3 = new ParamInfo();
        blockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D", "BR2-S-DDD");
        blockInfo3.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BR2-S-DDD2");

        //BCブロック4データ
        ParamInfo blockInfo4 = new ParamInfo();
        blockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D", "BR2-S-DDD");
        blockInfo4.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BR2-S-DDD2");

        //BCブロック5データ
        ParamInfo blockInfo5 = new ParamInfo();
        blockInfo5.addParam(SingleParamParser.DEFAULT_TAG, "D", "BR5-S-DDD");
        blockInfo5.addParam(SingleParamParser.DEFAULT_TAG, "D2", "BR5-S-DDD2");
        
        for (int i = 0; i < reportSheets.length; i++) {
            ParamInfo info = reportSheets[i].getParamInfo();
            if (i == 1) {
                info.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "BR1", new Object[]{blockInfo1, blockInfo2, blockInfo3, blockInfo4, blockInfo5});
            } else {
                info.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "BR1", new Object[]{blockInfo1, blockInfo2});    
            }
        }

        BlockRowRepeatParamParser parser = new BlockRowRepeatParamParser();
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
        // BR-C
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet1 = workbook.getSheetAt( 0);
        results = parseSheet( parser, sheet1, reportsParserInfo);
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 6, 7)};
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
        // BR-C2
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet2 = workbook.getSheetAt( 1);
        
        results = parseSheet( parser, sheet2, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 6, 9)};
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
        // BR-R
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet3 = workbook.getSheetAt( 2);
        
        results = parseSheet( parser, sheet3, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 20, 3)};
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
        // BR-R2
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet4 = workbook.getSheetAt( 3);
        
        results = parseSheet( parser, sheet4, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 3, 3)};
        expectAfCells = new CellObject[] {new CellObject( 24, 3)};
        checkResult( expectBeCells, expectAfCells, results);
        
        checkSheet( "Sheet4", sheet4, true);
        
        
        // -----------------------
        // □[正常系]オプション指定なし
        // BR-CR
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet5 = workbook.getSheetAt( 4);
        
        results = parseSheet( parser, sheet5, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 4, 3)};
        expectAfCells = new CellObject[] {new CellObject( 15, 7)};
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
        // BR-BR
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet6 = workbook.getSheetAt( 5);
        
        results = parseSheet( parser, sheet6, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 8, 5), new CellObject( -1, -1), new CellObject( -1, -1)};
        expectAfCells = new CellObject[] {new CellObject( 31, 9), new CellObject( -1, -1), new CellObject( -1, -1)};
        
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
        // BR-BC
        // -----------------------
        workbook = getWorkbook();
        Sheet sheet7 = workbook.getSheetAt( 6);
        
        results = parseSheet( parser, sheet7, reportsParserInfo);
        
        expectBeCells = new CellObject[] {new CellObject( 8, 5)};
        expectAfCells = new CellObject[] {new CellObject( 22, 16)};
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
        expectAfCells = new CellObject[] {new CellObject( 12, 3)};
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
            fail( "repeatNumのマイナスチェックにかかっていない");
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
        expectAfCells = new CellObject[] {new CellObject( 12, 3)};
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
        expectAfCells = new CellObject[] {new CellObject( 30, 15)};
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
        
        expectBeCells = new CellObject[] {new CellObject( 3, 4)};
        expectAfCells = new CellObject[] {new CellObject( 6, 4)};
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
        
        expectBeCells = new CellObject[] {new CellObject( 3, 4)};
        expectAfCells = new CellObject[] {new CellObject( 9, 4)};
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
     * {@link org.bbreak.excella.reports.tag.BlockRowRepeatParamParser#useControlRow()} のためのテスト・メソッド。
     */
    @Test
    public void testUseControlRow() {
        // -----------------------
        // □制御行使用有無：無
        // -----------------------
        BlockRowRepeatParamParser parser = new BlockRowRepeatParamParser();
        assertTrue( parser.useControlRow());
    }
    
    
    /**
     * {@link org.bbreak.excella.reports.tag.BlockRowRepeatParamParser#BlockColRepeatParamParser(java.lang.String)} のためのテスト・メソッド。
     */
    @Test
    public void testBlockRowRepeatParamParserString() {

        BlockRowRepeatParamParser parser = new BlockRowRepeatParamParser( "てすと");

        assertEquals( "てすと", parser.getTag());
    }
    
    private void checkSheet( String expectedSheetName, Sheet actualSheet, boolean outputExcel) {
        Workbook expectedWorkbook = getExpectedWorkbook();
        checkSheet( expectedWorkbook, expectedSheetName, actualSheet, outputExcel);
    }

    private void checkSheet( Workbook expectedWorkbook, String expectedSheetName, Sheet actualSheet, boolean outputExcel) {

        // 期待値ブックの読み込み
        Sheet expectedSheet = expectedWorkbook.getSheet( expectedSheetName);

        try {
            // チェック
            ReportsTestUtil.checkSheet( expectedSheet, actualSheet, false);
        } catch ( ReportsCheckException e) {
            fail( e.getCheckMessagesToString());
        } finally {
            String tmpDirPath = ReportsTestUtil.getTestOutputDir();
            try {
                Date now = new Date();
                String filepath = tmpDirPath + this.getClass().getSimpleName() + now.getTime() + getSuffix();
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

    /**
     * Test case for <a href="https://github.com/excella-core/excella-reports/issues/57">Issue #57</a>.
     */
    @Test
    public void testRepeatFormulas() throws ParseException, EncryptedDocumentException, IOException {
        List<ParamInfo> params = new ArrayList<>();
        for ( int i = 0; i < 3; i++) {
            ParamInfo param = new ParamInfo();
            param.addParam( SingleParamParser.DEFAULT_TAG, "col1", BigDecimal.valueOf( 1).movePointRight( i));
            param.addParam( SingleParamParser.DEFAULT_TAG, "col2", BigDecimal.valueOf( 2).movePointRight( i));
            param.addParam( SingleParamParser.DEFAULT_TAG, "col3", BigDecimal.valueOf( 3).movePointRight( i));
            params.add( param);
        }
        Workbook workbook = WorkbookFactory.create( getClass().getResourceAsStream( "BlockRowRepeatParamParser_formulas" + getSuffix()));

        ReportBook reportBook = new ReportBook( "", "test", new ConvertConfiguration[] {});
        ReportSheet reportSheet = new ReportSheet( "issue57");
        reportSheet.addParam( BlockRowRepeatParamParser.DEFAULT_TAG, "data", params.toArray());

        Sheet sheet = workbook.getSheet( "issue-57");
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.setReportParsers( new ArrayList<ReportsTagParser<?>>( ReportCreateHelper.createDefaultParsers().values()));
        reportsParserInfo.setReportBook( reportBook);
        reportsParserInfo.setParamInfo( reportSheet.getParamInfo());
        BlockRowRepeatParamParser parser = new BlockRowRepeatParamParser();

        parseSheet( parser, sheet, reportsParserInfo);
        
        Workbook expectedWorkbook = WorkbookFactory.create( //
            getClass().getResourceAsStream( "BlockRowRepeatParamParser_formulas_expected" + getSuffix()));
        checkSheet( expectedWorkbook, "issue-57", sheet, true);
    }

}
