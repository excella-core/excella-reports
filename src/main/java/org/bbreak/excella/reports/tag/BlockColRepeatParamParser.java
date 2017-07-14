/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: BlockColRepeatParamParser.java 189 2010-08-17 18:22:43Z ogiharasf $
 * $Revision: 189 $
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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.util.ReportsUtil;

/**
 * シート内の置換文字列を横方向にブロック単位で繰り返し変換するパーサ
 * 
 * @since 1.0
 */
public class BlockColRepeatParamParser extends ReportsTagParser<Object[]> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( BlockColRepeatParamParser.class);

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$BC[]";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * ブロックリピート開始セルパラメータ
     */
    public static final String PARAM_FROM = "fromCell";

    /**
     * ブロックリピート終了セルパラメータ
     */
    public static final String PARAM_TO = "toCell";

    /**
     * 最大繰り返し回数パラメータ
     */
    public static final String PARAM_REPEAT = "repeatNum";

    /**
     * 最大折り返し回数パラメータ
     */
    public static final String PARAM_TURN = "turnNum";

    /**
     * 重複非表示の調整パラメータ
     */
    protected static final String PARAM_DUPLICATE = "hideDuplicate";

    /**
     * シートへのハイパーリンク設定有無
     */
    // 無しの方向
    // public static final String PARAM_SHEET_LINK = "sheetLink";
    /**
     * タグ除去パラメータ
     */
    public static final String PARAM_REMOVE_TAG = "removeTag";

    /**
     * コンストラクタ
     */
    public BlockColRepeatParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public BlockColRepeatParamParser( String tag) {
        super( tag);
    }

    @Override
    public boolean useControlRow() {
        return true;
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
        try {
            // パラメータの取得
            Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

            // パラメータチェック
            checkParam( paramDef, tagCell);

            ReportsParserInfo info = ( ReportsParserInfo) data;
            ParamInfo paramInfo = info.getParamInfo();

            // リターンクラス
            ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
            List<Object> resultList = new ArrayList<Object>();

            // BCタグの名前取得
            String bcTagname = paramDef.get( PARAM_VALUE);
            if ( log.isDebugEnabled()) {
                log.debug( "BCパラメータ名 : " + bcTagname);
            }

            // パラメータ名に対応するデータ取得
            Object[] paramInfos = getParamData( paramInfo, bcTagname);
            if ( paramInfos == null) {
                return parsedReportInfo;
            }
            // 単純置換のパーサー取得
            List<SingleParamParser> singleParsers = getSingleReplaceParsers( info);

            // POJOをParamInfoに変換する
            List<ParamInfo> paramInfoList = new ArrayList<ParamInfo>();
            for ( Object obj : paramInfos) {
                if ( obj instanceof ParamInfo) {
                    paramInfoList.add( ( ParamInfo) obj);
                    continue;
                }
                ParamInfo childParamInfo = new ParamInfo();
                Map<String, Object> map = PropertyUtils.describe( obj);
                for ( Map.Entry<String, Object> entry : map.entrySet()) {
                    for ( ReportsTagParser<?> parser : singleParsers) {
                        childParamInfo.addParam( parser.getTag(), entry.getKey(), entry.getValue());
                    }
                }
                paramInfoList.add( childParamInfo);
            }
            paramInfos = paramInfoList.toArray( new ParamInfo[paramInfoList.size()]);

            // 繰り返し最大回数
            Integer repeatNum = paramInfos.length;
            if ( paramDef.containsKey( PARAM_REPEAT)) {
                if ( Integer.valueOf( paramDef.get( PARAM_REPEAT)) < repeatNum) {
                    repeatNum = Integer.valueOf( paramDef.get( PARAM_REPEAT));
                }
            }

            // 折り返し最大回数
            // TODO 保留
            /*
             * Integer turnNum = null; if (paramDef.containsKey( PARAM_TURN)) { turnNum = Integer.valueOf( paramDef.get( PARAM_TURN)); }
             */

            // 重複表示不可のためのシングル置換結果格納マップ
            Map<String, Object> beforeBlockSingleDataMap = new HashMap<String, Object>();
            // 重複非表示の設定有無
            if ( paramDef.containsKey( PARAM_DUPLICATE)) {
                String[] tmp = paramDef.get( PARAM_DUPLICATE).split( ";");
                // singleタグ形式にする
                for ( String str : tmp) {
                    for ( ReportsTagParser<?> parser : singleParsers) {
                        str = parser.getTag() + TAG_PARAM_PREFIX + str + TAG_PARAM_SUFFIX;
                        // 最初はkeyだけセットする
                        beforeBlockSingleDataMap.put( str, "");
                    }
                }
            }

            // removeTag定義
            boolean removeTag = false;
            if ( paramDef.containsKey( PARAM_REMOVE_TAG)) {
                removeTag = Boolean.valueOf( paramDef.get( PARAM_REMOVE_TAG));
            }

            // fromCell定義
            int[] fromCellPosition = ReportsUtil.getCellIndex( paramDef.get( PARAM_FROM), PARAM_FROM);
            int defaultFromCellRowIndex = tagCell.getRow().getRowNum() + fromCellPosition[0];
            int defaultFromCellColIndex = tagCell.getColumnIndex() + fromCellPosition[1];

            // toCell定義
            int[] toCellIndex = ReportsUtil.getCellIndex( paramDef.get( PARAM_TO), PARAM_TO);
            int defaultToCellRowIndex = tagCell.getRow().getRowNum() + toCellIndex[0];
            int defaultToCellColIndex = tagCell.getColumnIndex() + toCellIndex[1];

            // ブロック置換後最終座標(初期値はブロック範囲)
            int blockEndRowIndex = defaultToCellRowIndex;
            int blockEndColIndex = defaultToCellColIndex;
            int blockStartRowIndex = defaultFromCellRowIndex;
            int blockStartColIndex = defaultFromCellColIndex;

            // BCブロック情報の保存Cell[row][col]
            Object[][] blockCellsValue = ReportsUtil.getBlockCellValue( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            CellStyle[][] blockCellsStyle = ReportsUtil.getBlockCellStyle( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            CellType[][] blockCellTypes = ReportsUtil.getBlockCellType( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            // 最終列座標までの列幅情報を保存する
            int[] columnWidths = ReportsUtil.getColumnWidth( sheet, defaultFromCellColIndex, defaultToCellColIndex);

            // ブロック範囲取得
            int rowlen = defaultToCellRowIndex - defaultFromCellRowIndex + 1;
            int collen = defaultToCellColIndex - defaultFromCellColIndex + 1;

            // 対象シートより結合セルを取得
            CellRangeAddress[] margedCells = ReportsUtil.getMargedCells( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);

            // 折り返しまでのブロックループカウンタ
            int turnCount = 0;

            int maxblockEndRowIndex = blockEndRowIndex;

            TagParser<?> parser = null;

            ParsedReportInfo result = null;

            // データループ
            for ( int repeatCount = 0; repeatCount < repeatNum; repeatCount++) {

                collen = defaultToCellColIndex - defaultFromCellColIndex + 1;

                // データループが折り返し回数に達したらblockStartRowIndexの設定、blockEndColIndexのリセット
                // TODO 保留
                /*
                 * if (turnNum != null && turnNum <= turnCount) { blockStartRowIndex = blockEndRowIndex + 1; blockEndColIndex = 0; turnCount = 0; }
                 */

                // ブロック開始列設定
                if ( turnCount > 0) {
                    blockStartColIndex = blockEndColIndex + 1;
                    blockEndColIndex = blockStartColIndex + collen - 1;
                } else {
                    // スタート列は、指定された列なのでコメントアウト 2009/11/16
                    // blockStartColIndex = tagCell.getColumnIndex();
                }
                turnCount++;

                // 右方向にBCブロック分
                if ( repeatCount > 0) {
                    CellRangeAddress rangeAddress = new CellRangeAddress( blockStartRowIndex, sheet.getLastRowNum(), blockStartColIndex, blockStartColIndex + collen - 1);
                    PoiUtil.insertRangeRight( sheet, rangeAddress);
                    if ( log.isDebugEnabled()) {
                        log.debug( "デフォルトブロック挿入");
                        log.debug( blockStartRowIndex + ":" + sheet.getLastRowNum() + ":" + blockStartColIndex + ":" + (blockStartColIndex + collen - 1));
                    }

                    // ■結合セルの設定
                    // 編集処理中シートはリアルタイムで更新されていくため、
                    // 結合セルの指定には動的補正を行う
                    int targetColNum = blockEndColIndex - defaultToCellColIndex;
                    for ( CellRangeAddress address : margedCells) {
                        // 開始行＝結合セル開始行(※BCタグは列方向に展開されるため動的補正不要)
                        // 終了行＝結合セル終了行(※BCタグは列方向に展開されるため動的補正不要)
                        // 開始列＝結合セル開始列 + BCタグ現在処理済最終列 - BCタグ存在位置
                        // 終了列＝結合セル終了列 + BCタグ現在処理済最終列 - BCタグ存在位置
                        int firstRowNum = address.getFirstRow();
                        int lastRowNum = address.getLastRow();
                        int firstColumnNum = address.getFirstColumn() + targetColNum;
                        int lastColumnNum = address.getLastColumn() + targetColNum;

                        CellRangeAddress copyAddress = new CellRangeAddress( firstRowNum, lastRowNum, firstColumnNum, lastColumnNum);
                        sheet.addMergedRegion( copyAddress);
                    }

                }

                // ブロック情報をシートにコピー
                if ( log.isDebugEnabled()) {
                    log.debug( "●●●ブロック情報をコピー●●● データループ=" + repeatCount);
                }

                for ( int rowIdx = 0; rowIdx < blockCellsValue.length; rowIdx++) {
                    // ■行取得
                    // コピー先行インデックス
                    int copyToRowIndex = blockStartRowIndex + rowIdx;
                    Row row = sheet.getRow( copyToRowIndex);
                    // コピー先の行がnullかつ、コピー元の行に何らかの情報を持っている場合は行を生成する
                    if ( row == null && !ReportsUtil.isEmptyRow( blockCellTypes[rowIdx], blockCellsValue[rowIdx], blockCellsStyle[rowIdx])) {
                        // カバレージが通らないデッドロジック
                        // ※理由
                        // (行範囲は固定で)列方向に範囲展開していくことから
                        // 展開先の行がnullである時に(RowをCreateする)セル情報の設定をするという事象が起こり得ないため
                        row = sheet.createRow( copyToRowIndex);
                    }

                    if ( row != null) {
                        // ■セル取得＆設定
                        for ( int colIdx = 0; colIdx < blockCellsValue[rowIdx].length; colIdx++) {
                            // コピー先列インデックス
                            int copyToColIndex = blockStartColIndex + colIdx;
                            // 列幅の設定
                            sheet.setColumnWidth( copyToColIndex, columnWidths[colIdx]);
                            // セル取得
                            Cell cell = row.getCell( copyToColIndex);
                            // セルタイプの取得
                            CellType cellType = blockCellTypes[rowIdx][colIdx];
                            // セル値の取得
                            Object cellValue = blockCellsValue[rowIdx][colIdx];
                            // セルスタイルの取得
                            CellStyle cellStyle = blockCellsStyle[rowIdx][colIdx];
                            // セルに設定すべき情報（タイプ、値、スタイルのいずれか）が有る場合のみ生成する
                            if ( cell == null && !ReportsUtil.isEmptyCell( cellType, cellValue, cellStyle)) {
                                cell = row.createCell( copyToColIndex);
                            }

                            // セル設定
                            if ( cell != null) {
                                // セルタイプの設定
                                cell.setCellType( cellType);
                                // セルの値の設定
                                PoiUtil.setCellValue( cell, cellValue);
                                // セルのスタイルの設定
                                if ( cellStyle != null) {
                                    cell.setCellStyle( cellStyle);
                                }
                                if ( log.isDebugEnabled()) {
                                    log.debug( "row=" + (copyToRowIndex) + " col" + (copyToColIndex) + ">>>>>>" + blockCellsValue[rowIdx][colIdx]);
                                }
                            }
                        }
                    }
                }

                // 子パーサでの行挿入によって親タグ解析範囲の増加の値
                int plusRowNum = 0;
                // 子パーサでの列挿入によって親タグ解析範囲の増加の値
                int plusColNum = 0;
                collen = defaultToCellColIndex - defaultFromCellColIndex + 1;

                // 行ループ
                for ( int targetRow = blockStartRowIndex; targetRow < blockStartRowIndex + rowlen + plusRowNum; targetRow++) {
                    if ( sheet.getRow( targetRow) == null) {
                        if ( log.isDebugEnabled()) {
                            log.debug( "row=" + targetRow + " : row is not available. continued...");
                        }
                        continue;
                    }

                    // 置換対象セル
                    Cell chgTargetCell = null;

                    // 列ループ
                    for ( int targetCol = blockStartColIndex; targetCol <= blockStartColIndex + collen + plusColNum - 1; targetCol++) {

                        // 置換されるセルの取得
                        chgTargetCell = sheet.getRow( targetRow).getCell( targetCol);
                        if ( chgTargetCell == null) {
                            if ( log.isDebugEnabled()) {
                                log.debug( "row=" + targetRow + " col=" + targetCol + " : cell is not available. continued...");
                            }
                            continue;
                        }

                        parser = info.getMatchTagParser( sheet, chgTargetCell);
                        if ( parser == null) {
                            if ( log.isDebugEnabled()) {
                                log.debug( "row=" + targetRow + " col=" + targetCol + " parser is not available. continued...");
                            }
                            continue;
                        }

                        String chgTargetCellString = chgTargetCell.getStringCellValue();
                        if ( log.isDebugEnabled()) {
                            log.debug( "########## タグセル座標 row=" + targetRow + " col=" + targetCol + " タグ文字=" + chgTargetCellString + " ##########");
                        }

                        // パース処理実行
                        result = ( ParsedReportInfo) parser.parse( sheet, chgTargetCell, info.createChildParserInfo( ( ParamInfo) paramInfos[repeatCount]));

                        // 子パーサでの行挿入数
                        plusRowNum += result.getRowIndex() - result.getDefaultRowIndex();

                        // 子パーサでの列挿入数
                        plusColNum += result.getColumnIndex() - result.getDefaultColumnIndex();

                        // パーサーがsingle置換の場合オプションで指定された重複不可タグの置換値
                        // が前ブロックの同置換値と同じ場合は消去する。
                        if ( parser instanceof SingleParamParser && beforeBlockSingleDataMap.containsKey( chgTargetCellString)) {
                            if ( beforeBlockSingleDataMap.get( chgTargetCellString).equals( result.getParsedObject())) {
                                // 空文字セット
                                PoiUtil.setCellValue( chgTargetCell, "");

                            } else {
                                // マップの値の上書き
                                beforeBlockSingleDataMap.put( chgTargetCellString, result.getParsedObject());
                            }
                        }

                        // 子パーサによって縦に増えた分、子パーサのタグ位置より前の列にセルを挿入
                        if ( blockStartColIndex != result.getDefaultColumnIndex() && maxblockEndRowIndex <= blockEndRowIndex && result.getRowIndex() > result.getDefaultRowIndex()) {
                            CellRangeAddress preRangeAddress = new CellRangeAddress( blockEndRowIndex + 1, blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex()), blockStartColIndex,
                                targetCol - 1);
                            PoiUtil.insertRangeDown( sheet, preRangeAddress);
                            if ( log.isDebugEnabled()) {
                                log.debug( "***縦補正***");
                                log.debug( "挿入範囲1 : " + (blockEndRowIndex + 1) + ":" + (blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex())) + ":" + blockStartColIndex + ":"
                                        + (targetCol - 1));
                            }
                        }

                        // Rの場合はタグ位置以降にもセルを挿入
                        if ( parser instanceof RowRepeatParamParser && maxblockEndRowIndex <= blockEndRowIndex && result.getRowIndex() > result.getDefaultRowIndex()) {
                            CellRangeAddress rearRangeAddress = new CellRangeAddress( blockEndRowIndex + 1, blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex()), result
                                .getDefaultColumnIndex() + 1, blockEndColIndex);
                            PoiUtil.insertRangeDown( sheet, rearRangeAddress);
                            if ( log.isDebugEnabled()) {
                                log.debug( "***縦補正***");
                                log.debug( "挿入範囲2 : " + (blockEndRowIndex + 1) + ":" + (blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex())) + ":"
                                        + (result.getDefaultColumnIndex() + 1) + ":" + blockEndColIndex);
                            }
                        }

                        blockEndRowIndex = defaultToCellRowIndex + plusRowNum;

                        resultList.add( result.getParsedObject());

                        if ( parser instanceof BlockColRepeatParamParser || parser instanceof BlockRowRepeatParamParser) {
                            collen += result.getColumnIndex() - result.getDefaultColumnIndex();
                            plusColNum -= result.getColumnIndex() - result.getDefaultColumnIndex();
                        }

                        // 子のパースよって横に増えた分、ブロック最終列座標を更新
                        if ( blockStartColIndex + collen + plusColNum - 1 > blockEndColIndex) {
                            int beforeLastColIndex = blockEndColIndex;
                            blockEndColIndex = blockStartColIndex + collen + plusColNum - 1;

                            // タグセル行から現在行の１つ上までに横に増えたセルを挿入する
                            CellRangeAddress preRangeAddress = new CellRangeAddress( tagCell.getRowIndex(), targetRow - 1, beforeLastColIndex + 1, blockEndColIndex);
                            PoiUtil.insertRangeRight( sheet, preRangeAddress);
                            if ( log.isDebugEnabled()) {
                                log.debug( "***横補正***");
                                log.debug( "挿入範囲1 : " + tagCell.getRowIndex() + ":" + (targetRow - 1) + ":" + (beforeLastColIndex + 1) + ":" + blockEndColIndex);
                            }
                            // 処理中最終行に最終列以降に増えたセル数列挿入
                            int lastRowNum = sheet.getLastRowNum();
                            if ( (blockEndRowIndex + 1) <= lastRowNum && (beforeLastColIndex + 1) <= blockEndColIndex) {
                                CellRangeAddress rangeAddress = new CellRangeAddress( blockEndRowIndex + 1, lastRowNum, beforeLastColIndex + 1, blockEndColIndex);
                                PoiUtil.insertRangeRight( sheet, rangeAddress);
                                if ( log.isDebugEnabled()) {
                                    log.debug( "***横補正***");
                                    log.debug( "挿入範囲3 : " + (blockEndRowIndex + 1) + ":" + lastRowNum + ":" + (beforeLastColIndex + 1) + ":" + blockEndColIndex);
                                }
                            }
                        }
                        // 列ループ終わり
                    }
                    // 処理行に横に増えたセルを挿入する
                    if ( blockStartColIndex + collen + plusColNum - 1 < blockEndColIndex) {
                        CellRangeAddress rearRangeAddress = new CellRangeAddress( targetRow, targetRow, blockStartColIndex + collen + plusColNum, blockEndColIndex);
                        PoiUtil.insertRangeRight( sheet, rearRangeAddress);
                        if ( log.isDebugEnabled()) {
                            log.debug( "***横補正***");
                            log.debug( "挿入範囲2 : " + targetRow + ":" + targetRow + ":" + (blockStartColIndex + collen + plusColNum) + ":" + blockEndColIndex);
                        }
                    }

                    plusColNum = 0;

                    // 行ループ終わり
                }

                // ブロック開始列より前に縦に伸びた分のセルを挿入する
                if ( maxblockEndRowIndex < blockEndRowIndex) {
                    if ( log.isDebugEnabled()) {
                        log.debug( "***縦補正***");
                    }
                    if ( repeatCount != 0) {
                        CellRangeAddress preRangeAddress = new CellRangeAddress( maxblockEndRowIndex + 1, blockEndRowIndex, defaultFromCellColIndex, blockStartColIndex - 1);
                        PoiUtil.insertRangeDown( sheet, preRangeAddress);
                        if ( log.isDebugEnabled()) {
                            log.debug( "挿入範囲1 : " + (maxblockEndRowIndex + 1) + ":" + blockEndRowIndex + ":" + defaultFromCellColIndex + ":" + (blockStartColIndex - 1));
                        }
                    }

                    // ブロック最終列以降に縦に伸びた分のセルを挿入する
                    CellRangeAddress rearRangeAddress = new CellRangeAddress( maxblockEndRowIndex + 1, blockEndRowIndex, blockEndColIndex + 1, PoiUtil.getLastColNum( sheet));
                    PoiUtil.insertRangeDown( sheet, rearRangeAddress);
                    if ( log.isDebugEnabled()) {
                        log.debug( "挿入範囲2 : " + (maxblockEndRowIndex + 1) + ":" + blockEndRowIndex + ":" + (blockEndColIndex + 1) + ":" + PoiUtil.getLastColNum( sheet));
                    }

                    maxblockEndRowIndex = blockEndRowIndex;
                }
            }

            // タグ除去
            if ( removeTag) {
                tagCell.setCellType( CellType.BLANK);
            }

            // リターン情報セット
            parsedReportInfo.setDefaultRowIndex( defaultToCellRowIndex);
            parsedReportInfo.setDefaultColumnIndex( defaultToCellColIndex);
            parsedReportInfo.setColumnIndex( blockEndColIndex);
            parsedReportInfo.setParsedObject( resultList);
            parsedReportInfo.setRowIndex( maxblockEndRowIndex);

            if ( log.isDebugEnabled()) {
                log.debug( "finalBlockRowIndex= " + maxblockEndRowIndex);
                log.debug( "finalBlockColIndex=" + blockEndColIndex);
            }
            return parsedReportInfo;

        } catch ( Exception e) {
            throw new ParseException( tagCell, e);
        }
    }

    /**
     * 現在使用されている単純置換パーサーを取得する。
     * 
     * @param reportsParserInfo 帳票解析処理情報
     * @return 現在使用されている単純置換
     */
    private List<SingleParamParser> getSingleReplaceParsers( ReportsParserInfo reportsParserInfo) {
        List<ReportsTagParser<?>> parsers = reportsParserInfo.getReportParsers();
        List<SingleParamParser> singleParsers = new ArrayList<SingleParamParser>();
        for ( ReportsTagParser<?> reportsParser : parsers) {
            if ( reportsParser instanceof SingleParamParser) {
                singleParsers.add( ( SingleParamParser) reportsParser);
            }
        }
        return singleParsers;
    }

    /**
     * 不正なパラメータがある場合、ParseExceptionをthrowする。
     * 
     * @param paramDef
     * @param tagCell
     * @throws ParseException
     */
    private void checkParam( Map<String, String> paramDef, Cell tagCell) throws ParseException {
        // キー
        // 必須チェック
        if ( !paramDef.containsKey( PARAM_FROM)) {
            throw new ParseException( tagCell + "必須パラメータなし：" + PARAM_FROM);
        }
        if ( !paramDef.containsKey( PARAM_TO)) {
            throw new ParseException( tagCell + "必須パラメータなし：" + PARAM_TO);
        }

        // 値
        if ( paramDef.get( PARAM_FROM).split( ":").length != 2) {
            throw new ParseException( tagCell + "パラメータ値不正：" + PARAM_FROM);
        }
        if ( paramDef.get( PARAM_TO).split( ":").length != 2) {
            throw new ParseException( tagCell + "パラメータ値不正：" + PARAM_TO);
        }

        int[] fromCellIndex = ReportsUtil.getCellIndex( paramDef.get( PARAM_FROM), PARAM_FROM);
        int[] toCellIndex = ReportsUtil.getCellIndex( paramDef.get( PARAM_TO), PARAM_TO);

        if ( fromCellIndex[0] < 0 || fromCellIndex[1] < 0) {
            throw new ParseException( tagCell + "パラメータ不正：" + PARAM_FROM);
        }
        if ( toCellIndex[0] < 0 || toCellIndex[1] < 0) {
            throw new ParseException( tagCell + "パラメータ不正：" + PARAM_TO);
        }

        if ( fromCellIndex[0] > toCellIndex[0]) {
            throw new ParseException( tagCell + "パラメータ不正（行）：" + PARAM_FROM + "," + PARAM_TO);
        }
        if ( fromCellIndex[1] > toCellIndex[1]) {
            throw new ParseException( tagCell + "パラメータ不正（列）：" + PARAM_FROM + "," + PARAM_TO);
        }

        try {
            if ( paramDef.containsKey( PARAM_REPEAT)) {
                if ( Integer.valueOf( paramDef.get( PARAM_REPEAT)) < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_REPEAT);
                }
            }
        } catch ( NumberFormatException e) {
            throw new ParseException( tagCell + "パラメータ不正：" + PARAM_REPEAT);
        }

    }
}
