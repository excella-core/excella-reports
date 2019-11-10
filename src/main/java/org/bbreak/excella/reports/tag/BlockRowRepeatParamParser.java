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
 * 指定された範囲を縦方向に繰り返して置換するパーサ
 * 
 * @since 1.0
 */
public class BlockRowRepeatParamParser extends ReportsTagParser<Object[]> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( BlockRowRepeatParamParser.class);

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$BR[]";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * 範囲fromパラメータ
     */
    public static final String PARAM_FROM_CELL = "fromCell";

    /**
     * 範囲toパラメータ
     */
    public static final String PARAM_TO_CELL = "toCell";

    /**
     * 繰り返し回数パラメータ
     */
    public static final String PARAM_REPEAT_NUM = "repeatNum";

    /**
     * 最小繰り返し回数パラメータ
     */
    public static final String PARAM_MIN_REPEAT_NUM = "minRepeatNum";
    
    /**
     * 折り返し回数パラメータ
     */
    public static final String PARAM_TURN_NUM = "turnNum";

    /**
     * 重複非表示パラメータ
     */
    public static final String PARAM_DUPLICATE = "hideDuplicate";

    /**
     * タグ除去パラメータ
     */
    public static final String PARAM_REMOVE_TAG = "removeTag";

    /**
     * コンストラクタ
     */
    public BlockRowRepeatParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public BlockRowRepeatParamParser( String tag) {
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
            checkParam( sheet, paramDef, tagCell);

            ReportsParserInfo reportsParserInfo = ( ReportsParserInfo) data;
            ParamInfo paramInfo = reportsParserInfo.getParamInfo();

            // parse結果オブジェクト
            ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
            List<Object> resultList = new ArrayList<Object>();

            // ブロック置換後最終座標
            int finalBlockRowIndex = 0;
            int finalBlockColIndex = 0;

            String brTagName = paramDef.get( PARAM_VALUE);
            if ( log.isDebugEnabled()) {
                log.debug( "BRパラメータ名: " + brTagName);
            }

            // パラメータ名に対応するデータ取得
            Object[] paramInfos = getParamData( paramInfo, brTagName);
            if ( paramInfos == null) {
                return parsedReportInfo;
            }

            // 単純置換のパーサー取得
            List<SingleParamParser> singleParsers = getSingleReplaceParsers( reportsParserInfo);

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
            
            // 最小繰り返し回数
            if ( paramDef.containsKey( PARAM_MIN_REPEAT_NUM)) {
                Integer minRepeatNum = Integer.valueOf( paramDef.get( PARAM_MIN_REPEAT_NUM));
                if ( minRepeatNum > paramInfoList.size()) {
                    int addEmptyRowNum = minRepeatNum - paramInfoList.size();
                    for ( int num = 0; num < addEmptyRowNum; num++) {
                        ParamInfo childParamInfo = new ParamInfo();
                        paramInfoList.add( childParamInfo);
                    }
                }
            }
            
            paramInfos = paramInfoList.toArray( new ParamInfo[paramInfoList.size()]);

            // repeatNum定義
            Integer repeatNum = paramInfos.length;
            if ( paramDef.containsKey( PARAM_REPEAT_NUM)) {
                if ( Integer.valueOf( paramDef.get( PARAM_REPEAT_NUM)) < repeatNum) {
                    repeatNum = Integer.valueOf( paramDef.get( PARAM_REPEAT_NUM));
                }
            }

            // duplicateParams定義 → 単純置換タグ形式としてエントリのみを追加
            Map<String, Object> unduplicableParamMap = new HashMap<String, Object>();
            if ( paramDef.containsKey( PARAM_DUPLICATE)) {
                String[] params = paramDef.get( PARAM_DUPLICATE).split( ";");
                for ( String param : params) {
                    for ( ReportsTagParser<?> parser : singleParsers) {
                        param = parser.getTag() + TAG_PARAM_PREFIX + param + TAG_PARAM_SUFFIX;
                        unduplicableParamMap.put( param, "");
                    }
                }
            }

            // removeTag定義
            boolean removeTag = false;
            if ( paramDef.containsKey( PARAM_REMOVE_TAG)) {
                removeTag = Boolean.valueOf( paramDef.get( PARAM_REMOVE_TAG));
            }

            // タグセル座標
            int tagCellRowIndex = tagCell.getRowIndex();
            int tagCellColIndex = tagCell.getColumnIndex();

            // fromCell定義
            String fromCellParamDef = paramDef.get( PARAM_FROM_CELL);
            int[] fromCellPosition = ReportsUtil.getCellIndex( fromCellParamDef, PARAM_FROM_CELL);
            int defaultFromCellRowIndex = tagCellRowIndex + fromCellPosition[0];
            int defaultFromCellColIndex = tagCellColIndex + fromCellPosition[1];

            // toCell定義
            String toCellParamDef = paramDef.get( PARAM_TO_CELL);
            int[] toCellPosition = ReportsUtil.getCellIndex( toCellParamDef, PARAM_TO_CELL);
            int defaultToCellRowIndex = tagCellRowIndex + toCellPosition[0];
            int defaultToCellColIndex = tagCellColIndex + toCellPosition[1];

            // セル値・セルスタイル定義
            Object[][] blockCellValues = ReportsUtil.getBlockCellValue( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            CellStyle[][] blockCellStyles = ReportsUtil.getBlockCellStyle( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            CellType[][] blockCellTypes = ReportsUtil.getBlockCellType( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);
            float[] rowHeight = ReportsUtil.getRowHeight( sheet, defaultFromCellRowIndex, defaultToCellRowIndex);

            // 対象シートより結合セルを取得
            CellRangeAddress[] margedCells = ReportsUtil.getMargedCells( sheet, defaultFromCellRowIndex, defaultToCellRowIndex, defaultFromCellColIndex, defaultToCellColIndex);

            // ブロックサイズ取得
            final int defaultBlockHeight = defaultToCellRowIndex - defaultFromCellRowIndex + 1;
            final int defaultBlockWidth = defaultToCellColIndex - defaultFromCellColIndex + 1;
            int rowlen = defaultBlockHeight;
            int collen = defaultBlockWidth;

            // 解析位置(fromCell定義位置から開始, 1ブロック解析後に位置の補正を行う)
            int blockStartRowIndex = defaultFromCellRowIndex;
            int blockStartColIndex = defaultFromCellColIndex;
            int blockEndRowIndex = defaultToCellRowIndex;
            int blockEndColIndex = defaultToCellColIndex;
            int maxblockEndRowIndex = blockEndRowIndex;

            parsedReportInfo.setDefaultRowIndex( defaultToCellRowIndex);
            parsedReportInfo.setDefaultColumnIndex( defaultToCellColIndex);

            for ( int repeatCount = 0; repeatCount < repeatNum; repeatCount++) {
                // ２回目以降はブロック定義範囲のセル挿入を行う
                if ( repeatCount > 0) {
                    blockStartRowIndex = blockEndRowIndex + 1;
                    blockEndRowIndex = blockStartRowIndex + rowlen - 1;

                    CellRangeAddress rangeAddress = new CellRangeAddress( blockStartRowIndex, blockEndRowIndex, blockStartColIndex, PoiUtil.getLastColNum( sheet));
                    PoiUtil.insertRangeDown( sheet, rangeAddress);

                    if ( log.isDebugEnabled()) {
                        log.debug( "テンプレート挿入");
                        log.debug( "挿入範囲 : " + blockStartRowIndex + ":" + (blockStartRowIndex + rowlen - 1) + ":" + blockStartColIndex + ":" + PoiUtil.getLastColNum( sheet));
                    }

                    // ■結合セルの設定
                    // 編集処理中シートはリアルタイムで更新されていくため、
                    // 結合セルの指定には動的補正を行う
                    int targetRowNum = maxblockEndRowIndex - (defaultFromCellRowIndex - 1);
                    for ( CellRangeAddress address : margedCells) {
                        // 開始行＝結合セル開始行 + BRタグ現在処理済最終行 - BRタグ存在位置
                        // 終了行＝結合セル終了行 + BRタグ現在処理済最終行 - BRタグ存在位置
                        // 開始列＝結合セル開始列(※BRタグは行方向に展開されるため動的補正不要)
                        // 終了列＝結合セル終了列(※BRタグは行方向に展開されるため動的補正不要)
                        int firstRowNum = address.getFirstRow() + targetRowNum;
                        int lastRowNum = address.getLastRow() + targetRowNum;
                        int firstColumnNum = address.getFirstColumn();
                        int lastColumnNum = address.getLastColumn();

                        CellRangeAddress copyAddress = new CellRangeAddress( firstRowNum, lastRowNum, firstColumnNum, lastColumnNum);
                        sheet.addMergedRegion( copyAddress);
                    }

                }

                if ( log.isDebugEnabled()) {
                    log.debug( "repeatCount = " + repeatCount);
                    log.debug( "blockStartRowIndex = " + blockStartRowIndex);
                    log.debug( "blockStartColIndex = " + blockStartColIndex);
                }

                // ブロック範囲の値・スタイルを適用
                if ( log.isDebugEnabled()) {
                    log.debug( "●●●ブロック情報をコピー●●● データループ=" + repeatCount);
                }

                for ( int rowIdx = 0; rowIdx < defaultBlockHeight; rowIdx++) {
                    // ■行取得
                    Row row = sheet.getRow( blockStartRowIndex + rowIdx);
                    // コピー先の行がnullかつ、コピー元の行に何らかの情報を持っている場合は行を生成する
                    if ( row == null && !ReportsUtil.isEmptyRow( blockCellTypes[rowIdx], blockCellValues[rowIdx], blockCellStyles[rowIdx])) {
                        row = sheet.createRow( blockStartRowIndex + rowIdx);
                    }

                    if ( row != null) {
                        // 行の高さの設定
                        row.setHeightInPoints( rowHeight[rowIdx]);

                        // ■セル取得＆設定
                        for ( int colIdx = 0; colIdx < defaultBlockWidth; colIdx++) {
                            // セル取得
                            Cell cell = row.getCell( blockStartColIndex + colIdx);
                            // セルタイプの取得
                            CellType cellType = blockCellTypes[rowIdx][colIdx];
                            // セル値の取得
                            Object cellValue = blockCellValues[rowIdx][colIdx];
                            // セルスタイルの取得
                            CellStyle cellStyle = blockCellStyles[rowIdx][colIdx];
                            // セルに設定すべき情報（タイプ、値、スタイルのいずれか）が有る場合のみ生成する
                            if ( cell == null && !ReportsUtil.isEmptyCell( cellType, cellValue, cellStyle)) {
                                cell = row.createCell( blockStartColIndex + colIdx);
                            }

                            // セル設定
                            if ( cell != null) {
                                // セルタイプの設定
                                cell.setCellType( cellType);
                                // セルの値の設定
                                PoiUtil.setCellValue( cell, cellValue);
                                // セルのスタイルの設定
                                if ( cellStyle == null) {
                                    log.info( "Cell Style at [" + rowIdx + "," + colIdx + "] is not available. Skipping setCellValue()");
                                } else {
                                    cell.setCellStyle( cellStyle);
                                }
                                log.debug( "row=" + (blockStartRowIndex + rowIdx) + " col" + (blockStartColIndex + colIdx) + ">>>>>>" + blockCellValues[rowIdx][colIdx]);
                            }
                        }
                    }
                }

                int currentBlockHeight = rowlen;
                int currentBlockWidth = collen;
                // 子パーサでの行挿入によって親タグ解析範囲の増加の値
                int plusRowNum = 0;
                // 子パーサでの行挿入によって親タグ解析範囲の増加の値
                int plusColNum = 0;
                collen = defaultBlockWidth;
                // 行ループ
                // ブロック範囲の値の解析
                for ( int targetRow = blockStartRowIndex; targetRow < blockStartRowIndex + rowlen + plusRowNum; targetRow++) {
                    if ( finalBlockRowIndex < targetRow) {
                        finalBlockRowIndex = targetRow;
                    }
                    if ( sheet.getRow( targetRow) == null) {
                        if ( log.isDebugEnabled()) {
                            log.debug( "row=" + targetRow + " : row is not available. continued...");
                        }
                        continue;
                    }

                    for ( int targetCol = blockStartColIndex; targetCol <= blockStartColIndex + collen + plusColNum - 1; targetCol++) {
                        if ( finalBlockColIndex < targetCol) {
                            finalBlockColIndex = targetCol;
                        }
                        Cell targetCell = sheet.getRow( targetRow).getCell( targetCol);
                        if ( targetCell == null) {
                            if ( log.isDebugEnabled()) {
                                log.debug( "row=" + targetRow + " col=" + targetCol + " : cell is not available. continued...");
                            }
                            continue;
                        }

                        // タグパーサの取得
                        TagParser<?> parser = reportsParserInfo.getMatchTagParser( sheet, targetCell);
                        if ( parser == null) {
                            if ( log.isDebugEnabled()) {
                                log.debug( "row=" + targetRow + " col=" + targetCol + " parser is not available. continued...");
                            }
                            continue;
                        }

                        String targetCellTag = targetCell.getStringCellValue();
                        if ( log.isDebugEnabled()) {
                            log.debug( "########## タグセル座標 row=" + targetRow + " col=" + targetCol + " タグ文字=" + targetCellTag + " ##########");
                        }

                        // パース実行
                        ParsedReportInfo result = ( ParsedReportInfo) parser.parse( sheet, targetCell, reportsParserInfo.createChildParserInfo( ( ParamInfo) paramInfos[repeatCount]));
                        resultList.add( result.getParsedObject());

                        // 子パーサでの行挿入数
                        plusRowNum += result.getRowIndex() - result.getDefaultRowIndex();

                        // 子パーサでの列挿入数
                        plusColNum += result.getColumnIndex() - result.getDefaultColumnIndex();

                        int additionalHeight = result.getRowIndex() - result.getDefaultRowIndex();
                        int additionalWidth = result.getColumnIndex() - result.getDefaultColumnIndex();

                        // 子パーサから受け取った最終座標から解析範囲の補正
                        currentBlockHeight = currentBlockHeight + additionalHeight;
                        currentBlockWidth = currentBlockWidth + additionalWidth;

                        // 重複不可制御
                        if ( parser instanceof SingleParamParser) {
                            if ( unduplicableParamMap.containsKey( targetCellTag)) {
                                if ( unduplicableParamMap.get( targetCellTag).equals( result.getParsedObject())) {
                                    PoiUtil.setCellValue( targetCell, "");

                                } else {
                                    unduplicableParamMap.put( targetCellTag, result.getParsedObject());
                                }
                            }
                        }

                        // 子パーサによって縦に増えた分、子パーサのタグ位置より前の列にセルを挿入
                        if ( defaultFromCellColIndex != result.getDefaultColumnIndex() && result.getRowIndex() > result.getDefaultRowIndex() && blockStartColIndex < targetCol) {
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
                        if ( parser instanceof RowRepeatParamParser && result.getRowIndex() > result.getDefaultRowIndex()) {
                            CellRangeAddress rearRangeAddress = new CellRangeAddress( blockEndRowIndex + 1, blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex()), result
                                .getDefaultColumnIndex() + 1, PoiUtil.getLastColNum( sheet));
                            PoiUtil.insertRangeDown( sheet, rearRangeAddress);
                            if ( log.isDebugEnabled()) {
                                log.debug( "***縦補正***");
                                log.debug( "挿入範囲2 : " + (blockEndRowIndex + 1) + ":" + (blockEndRowIndex + (result.getRowIndex() - result.getDefaultRowIndex())) + ":"
                                        + (result.getDefaultColumnIndex() + 1) + ":" + PoiUtil.getLastColNum( sheet));
                            }
                        }

                        blockEndRowIndex += result.getRowIndex() - result.getDefaultRowIndex();

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
                // ブロック最終列以降に縦に伸びた分のセルを挿入する
                int lastColNum = PoiUtil.getLastColNum( sheet);
                if ( (maxblockEndRowIndex + 1) <= blockEndRowIndex && (blockEndColIndex + 1) <= lastColNum) {
                    CellRangeAddress rearRangeAddress = new CellRangeAddress( maxblockEndRowIndex + 1, blockEndRowIndex, blockEndColIndex + 1, lastColNum);
                    PoiUtil.insertRangeDown( sheet, rearRangeAddress);
                    if ( log.isDebugEnabled()) {
                        log.debug( "挿入範囲2 : " + (maxblockEndRowIndex + 1) + ":" + blockEndRowIndex + ":" + (blockEndColIndex + 1) + ":" + lastColNum);
                    }
                }
                maxblockEndRowIndex = blockEndRowIndex;
            }

            // タグ除去
            if ( removeTag) {
                tagCell.setBlank();
            }

            // 解析結果
            parsedReportInfo.setColumnIndex( finalBlockColIndex);
            parsedReportInfo.setRowIndex( finalBlockRowIndex);
            parsedReportInfo.setParsedObject( resultList);
            parsedReportInfo.setDefaultRowIndex( defaultToCellRowIndex);
            parsedReportInfo.setDefaultColumnIndex( defaultToCellColIndex);

            if ( log.isDebugEnabled()) {
                log.debug( "finalBlockRowIndex= " + finalBlockRowIndex);
                log.debug( "finalBlockColIndex=" + finalBlockColIndex);
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
    private void checkParam( Sheet sheet, Map<String, String> paramDef, Cell tagCell) throws ParseException {

        // 必須オプションチェック
        // fromCell
        if ( !paramDef.containsKey( PARAM_FROM_CELL)) {
            throw new ParseException( tagCell, "必須パラメータなし：" + PARAM_FROM_CELL);
        }
        // toCell
        if ( !paramDef.containsKey( PARAM_TO_CELL)) {
            throw new ParseException( tagCell, "必須パラメータなし：" + PARAM_TO_CELL);
        }

        // 値
        if ( paramDef.get( PARAM_FROM_CELL).split( ":").length != 2) {
            throw new ParseException( tagCell + "パラメータ値不正：" + PARAM_FROM_CELL);
        }
        if ( paramDef.get( PARAM_TO_CELL).split( ":").length != 2) {
            throw new ParseException( tagCell + "パラメータ値不正：" + PARAM_TO_CELL);
        }

        // 値妥当性チェック
        // 範囲from、範囲to
        int[] fromCellIndex = ReportsUtil.getCellIndex( paramDef.get( PARAM_FROM_CELL), PARAM_FROM_CELL);
        int[] toCellIndex = ReportsUtil.getCellIndex( paramDef.get( PARAM_TO_CELL), PARAM_TO_CELL);
        // マイナスチェック
        if ( fromCellIndex[0] < 0 || fromCellIndex[1] < 0) {
            throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_FROM_CELL);
        }
        if ( toCellIndex[0] < 0 || toCellIndex[1] < 0) {
            throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_TO_CELL);
        }
        // 行範囲指定チェック
        if ( fromCellIndex[0] > toCellIndex[0]) {
            throw new ParseException( tagCell, "パラメータ不正（行）：" + PARAM_FROM_CELL + "," + PARAM_TO_CELL);
        }
        // 列範囲指定チェック
        if ( fromCellIndex[1] > toCellIndex[1]) {
            throw new ParseException( tagCell, "パラメータ不正（列）：" + PARAM_FROM_CELL + "," + PARAM_TO_CELL);
        }

        // 繰り返し回数チェック（数値）
        if ( paramDef.containsKey( PARAM_REPEAT_NUM)) {
            try {
                if ( Integer.valueOf( paramDef.get( PARAM_REPEAT_NUM)) < 0) {
                    throw new ParseException( tagCell, "パラメータ値不正：" + PARAM_REPEAT_NUM);
                }
            } catch ( NumberFormatException e) {
                throw new ParseException( tagCell + "パラメータ不正：" + PARAM_REPEAT_NUM);
            }
        }
    }
}
