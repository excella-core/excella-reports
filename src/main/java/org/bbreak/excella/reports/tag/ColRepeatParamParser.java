/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ColRepeatParamParser.java 201 2013-03-15 05:11:47Z kamisono_bb $
 * $Revision: 201 $
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
import java.util.List;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bbreak.excella.core.CellClone;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.util.ReportsUtil;

/**
 * シート内の置換文字列を横方向に繰り返し変換するパーサ
 * 
 * @since 1.0
 */
public class ColRepeatParamParser extends ReportsTagParser<Object[]> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ColRepeatParamParser.class);

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$C[]";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * 重複非表示の調整パラメータ
     */
    protected static final String PARAM_DUPLICATE = "hideDuplicate";

    /**
     * 繰り返し最大回数
     */
    public static final String PARAM_REPEAT_NUM = "repeatNum";

    /**
     * シートへのハイパーリンク設定有無
     */
    public static final String PARAM_SHEET_LINK = "sheetLink";

    /**
     * シート変数の変数名
     */
    public static final String PARAM_PROPERTY = "property";
    
    /**
     * コンストラクタ
     */
    public ColRepeatParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ColRepeatParamParser( String tag) {
        super( tag);
    }

    @Override
    public boolean useControlRow() {
        return false;
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

        // パラメータチェック
        checkParam( paramDef, tagCell);

        String tag = tagCell.getStringCellValue();
        ReportsParserInfo info = ( ReportsParserInfo) data;
        ParamInfo paramInfo = info.getParamInfo();
        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();

        // 置換する値
        Object[] paramValues = null;
        try {
            // 置換変数の取得
            String replaceParam = paramDef.get( PARAM_VALUE);

            // 繰り返し最大回数
            Integer repeatNum = null;
            if ( paramDef.containsKey( PARAM_REPEAT_NUM)) {
                repeatNum = Integer.valueOf( paramDef.get( PARAM_REPEAT_NUM));
            }

            // シートハイパーリンク設定有無
            boolean sheetLink = false;
            if ( paramDef.containsKey( PARAM_SHEET_LINK)) {
                sheetLink = Boolean.valueOf( paramDef.get( PARAM_SHEET_LINK));
            }

            // シート変数
            String propertyName = null;
            if ( paramDef.containsKey( PARAM_PROPERTY)) {
                propertyName = paramDef.get( PARAM_PROPERTY);
            }

            // 重複非表示の設定有無
            boolean hideDuplicate = false;
            if ( paramDef.containsKey( PARAM_DUPLICATE)) {
                hideDuplicate = Boolean.valueOf( paramDef.get( PARAM_DUPLICATE));
            }
            
            // システム変数
            if ( ReportsUtil.VALUE_SHEET_NAMES.equals( replaceParam)) {
                // シート名
                paramValues = ReportsUtil.getSheetNames( info.getReportBook()).toArray();
            } else if ( ReportsUtil.VALUE_SHEET_VALUES.equals( replaceParam)) {
                // シート値
                paramValues = ReportsUtil.getSheetValues( info.getReportBook(), propertyName, info.getReportParsers()).toArray();
            } else {
                // 置換する値を取得
                if ( paramInfo != null) {
                    paramValues = getParamData( paramInfo, replaceParam);
                }
            }

            if ( paramValues == null || paramValues.length == 0) {
                // 空文字置換
                paramValues = new Object[] {null};
            }

            // パラメータ重複判定
            if ( hideDuplicate && paramValues.length > 1) {
                List<Object> paramValuesList = new ArrayList<Object>();
                for ( int i = 0; i <= paramValues.length - 1; i++) {
                    // 重複があったら空を追加
                    if ( !paramValuesList.contains( paramValues[i])) {
                        paramValuesList.add( paramValues[i]);
                    } else {
                        paramValuesList.add( null);
                    }
                }
                paramValues = paramValuesList.toArray();
            }

            // シフトセル数を取得
            int shiftNum = paramValues.length;
            // パラメタ数を取得
            int paramLength = paramValues.length;

            // 対象セルの行番号を取得
            int defaultFromCellRowIndex = tagCell.getRowIndex();
            // 対象セルの列番号を取得
            int defaultFromCellColIndex = tagCell.getColumnIndex();

            // 単位セル列サイズ
            int unitColSize = 1;

            // ■対象シートより結合セルを取得
            List<CellRangeAddress> maegedAddresses = new ArrayList<CellRangeAddress>();
            for ( int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress targetAddress = sheet.getMergedRegion( i);
                maegedAddresses.add( targetAddress);
            }

            // ■単位セル行サイズの設定
            if ( maegedAddresses.size() > 0) {
                // 対象シート上に結合セルが存在した場合には結合セル分処理を繰り返す
                for ( CellRangeAddress curMergedAdress : maegedAddresses) {
                    if ( defaultFromCellColIndex == curMergedAdress.getFirstColumn() && defaultFromCellRowIndex == curMergedAdress.getFirstRow()) {
                        // 対象セルが結合セル開始列番号＆開始行番号であった場合には
                        // 単位セル列サイズに結合セル列数を設定する
                        unitColSize = curMergedAdress.getLastColumn() - curMergedAdress.getFirstColumn() + 1;

                        // シフトセル数を単位セル列サイズ分（結合セル列数分）増加させる
                        shiftNum = shiftNum * unitColSize;
                    }
                }
            }

            // ■シフトセル数の設定
            if ( repeatNum != null && repeatNum < shiftNum) {
                // 繰り返し最大回数が指定されており、
                // かつ、シフトセル数が繰り返し最大回数を越えてしまっていた場合には
                // 繰り返し最大回数(repeatNum)×単位セル列サイズ(unitColSize)をシフトセル数とする
                shiftNum = repeatNum * unitColSize;

                // パラメタ数を繰り返し最大回数に置き換える
                paramLength = repeatNum;
            }

            // ■シフト前のシート（フォーマット）情報を保存
            // 最終列座標までのセル情報を保存する
            tagCell = new CellClone( tagCell);
            List<Cell> cellList = new ArrayList<Cell>();
            int defaultToOverCellColIndex = tagCell.getColumnIndex() + unitColSize;
            for ( int i = defaultFromCellColIndex; i < defaultToOverCellColIndex; i++) {
                Row targetCellRow = sheet.getRow( tagCell.getRowIndex());
                cellList.add( new CellClone(targetCellRow.getCell(i)));
            }
            
            // ■シフト処理            
            if ( shiftNum > 1) {
                // 値の数分シフト(１つであれば、シフトしない)
                int shiftColSize = tagCell.getColumnIndex() + shiftNum - unitColSize - 1;
                // セル単位でシフト
                CellRangeAddress rangeAddress = new CellRangeAddress( tagCell.getRowIndex(), tagCell.getRowIndex(), tagCell.getColumnIndex(), shiftColSize);
                PoiUtil.insertRangeRight( sheet, rangeAddress);
                // 列幅調整
                int tagCellWidth = sheet.getColumnWidth( tagCell.getColumnIndex());
                for ( int i = tagCell.getColumnIndex() + 1; i <= shiftColSize; i++) {
                    int colWidth = sheet.getColumnWidth( i);
                    if ( colWidth < tagCellWidth) {
                        // １つ後列の幅 ＜ 現在列の幅である場合には
                        // 現在列の幅を１つ後列の幅に設定する
                        sheet.setColumnWidth( i, tagCellWidth);
                    }
                }
            }

            // ■ブック、シートの取得
            Workbook workbook = sheet.getWorkbook();
            String sheetName = workbook.getSheetName( workbook.getSheetIndex( sheet));
            // シート名
            List<String> sheetNames = ReportsUtil.getSheetNames( info.getReportBook());
            // 変換値
            List<Object> resultValues = new ArrayList<Object>();
            // 前回配列値(beforeValue)
            Object beforeValue = null;

            // 設定値配列インデックス
            int valueIndex = -1;
            // シフトセル数分まで列毎の設定処理を行う
            for ( int colIndex = 0; colIndex < shiftNum; colIndex++) {
                // 行の取得
                Row row = sheet.getRow( tagCell.getRowIndex());
                if ( row == null) {
                    // カバレージが通らないデッドロジック
                    // ※理由
                    // (行は固定で)列方向に展開していくことから
                    // 展開先の行がnullである時に(RowをCreateする)セル情報の設定をするという事象が起こり得ないため
                    row = sheet.createRow( tagCell.getRowIndex());
                }
                // セルの取得
                Cell cell = row.getCell( tagCell.getColumnIndex() + colIndex);
                if ( cell == null) {
                    cell = row.createCell( tagCell.getColumnIndex() + colIndex);
                }
                // セルに設定する値の初期化(null)
                Object value = null;
                
                // セルの何番目に該当するかを除算にて算出
                // ※結果値が0の場合、(結合、非結合を問わず)最初のセルであることを示す
                int cellIndex = colIndex % unitColSize;

                // 結合行の値設定を行うセル以外は値を書き換える
                boolean skipCol = false;
                if ( cellIndex != 0) {
                    skipCol = true;
                } else {
                    valueIndex++;
                }

                // ■コピー
                // スタイル情報の設定
                PoiUtil.copyCell( cellList.get(cellIndex), cell);

                // 値の設定
                Object currentValue = paramValues[valueIndex];
                // 重複非表示オプション=trueかつ、値が一致する場合はフラグを書き換える
                boolean duplicateValue = false;
                if ( beforeValue != null && currentValue != null && beforeValue.equals( currentValue)) {
                    // 値が一致する場合
                    duplicateValue = true;
                }
                if ( !skipCol && !(hideDuplicate && duplicateValue)) {
                    // 重複非表示オプション=true
                    // かつ、値が一致する場合にはセルに設定する値を書き換える。
                    value = currentValue;
                }
                if ( log.isDebugEnabled()) {
                    log.debug( "[シート名=" + sheetName + ",セル=(" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")]  " + tag + " ⇒ " + value);
                }
                PoiUtil.setCellValue( cell, value);
                resultValues.add( value);

                // ■ハイパーリンク設定
                if ( sheetLink) {
                    if ( !skipCol && valueIndex < sheetNames.size()) {
                        PoiUtil.setHyperlink( cell, Hyperlink.LINK_DOCUMENT, "'" + sheetNames.get( valueIndex) + "'!A1");
                        if ( log.isDebugEnabled()) {
                            log.debug( "[シート名=" + sheetName + ",セル=(" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")]  Hyperlink ⇒ " + "'" + sheetNames.get( valueIndex) + "'!A1");
                        }
                    }
                }

                // ■結合セル指定
                // 元々テンプレートで結合済みのセルが存在するので最終の値の場合は結合しない
                if ( !skipCol && unitColSize > 1 && paramLength > valueIndex + 1) {
                    CellRangeAddress rangeAddress = new CellRangeAddress( cell.getRowIndex(), cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex() + unitColSize - 1);
                    sheet.addMergedRegion( rangeAddress);

                    // 結合の場合は結合したタイミングで前回の値を更新
                    beforeValue = value;
                }

                // 結合では無い場合は毎回値を更新
                if ( unitColSize == 1) {
                    beforeValue = value;
                }

            }

            parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex());
            // 最終列を設定（列結合セル状態も考慮）
            parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex() + unitColSize - 1);
            parsedReportInfo.setRowIndex( tagCell.getRowIndex());
            parsedReportInfo.setColumnIndex( tagCell.getColumnIndex() + shiftNum - 1);
            parsedReportInfo.setParsedObject( resultValues);
            if ( log.isDebugEnabled()) {
                log.debug( parsedReportInfo);
            }
            return parsedReportInfo;

        } catch ( Exception e) {
            throw new ParseException( tagCell, e);
        }

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
        if ( paramDef.containsKey( PARAM_DUPLICATE) && paramDef.containsKey( PARAM_SHEET_LINK)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_DUPLICATE + "," + PARAM_SHEET_LINK);
        }
    }

}
