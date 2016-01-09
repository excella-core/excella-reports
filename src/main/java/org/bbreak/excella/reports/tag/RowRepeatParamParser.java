/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: RowRepeatParamParser.java 201 2013-03-15 05:11:47Z kamisono_bb $
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
import org.apache.poi.common.usermodel.Hyperlink;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
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
 * シート内の行繰り返し置換文字列を変換するパーサ
 * 
 * @since 1.0
 */
public class RowRepeatParamParser extends ReportsTagParser<Object[]> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( RowRepeatParamParser.class);

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$R[]";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * 重複非表示の調整パラメータ
     */
    protected static final String PARAM_DUPLICATE = "hideDuplicate";

    /**
     * 値の挿入方法パラメータ
     */
    public static final String PARAM_ROW_SHIFT = "rowShift";

    /**
     * 改ページするデータ数
     */
    public static final String PARAM_BREAK_NUM = "breakNum";

    /**
     * 値変更での改ページ有無パラメータ
     */
    public static final String PARAM_CHANGE_BREAK = "changeBreak";

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
    public RowRepeatParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public RowRepeatParamParser( String tag) {
        super( tag);
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

        // パラメータチェック
        checkParam( paramDef, tagCell);

        String tag = tagCell.getStringCellValue();
        ReportsParserInfo reportsParserInfo = ( ReportsParserInfo) data;
        // 置換
        Object[] paramValues = null;
        try {
            // 値の挿入方法
            boolean rowShift = false;
            if ( paramDef.containsKey( PARAM_ROW_SHIFT)) {
                rowShift = Boolean.valueOf( paramDef.get( PARAM_ROW_SHIFT));
            }
            // 重複非表示
            boolean hideDuplicate = false;
            if ( paramDef.containsKey( PARAM_DUPLICATE)) {
                hideDuplicate = Boolean.valueOf( paramDef.get( PARAM_DUPLICATE));
            }
            // 改ページするデータ数
            Integer breakNum = null;
            if ( paramDef.containsKey( PARAM_BREAK_NUM)) {
                breakNum = Integer.valueOf( paramDef.get( PARAM_BREAK_NUM));
            }
            // 値変更改ページ有無
            boolean changeBreak = false;
            if ( paramDef.containsKey( PARAM_CHANGE_BREAK)) {
                changeBreak = Boolean.valueOf( paramDef.get( PARAM_CHANGE_BREAK));
            }
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
            // 置換変数の取得
            String replaceParam = paramDef.get( PARAM_VALUE);
            // システム変数
            if ( ReportsUtil.VALUE_SHEET_NAMES.equals( replaceParam)) {
                // シート名
                paramValues = ReportsUtil.getSheetNames( reportsParserInfo.getReportBook()).toArray();
            } else if ( ReportsUtil.VALUE_SHEET_VALUES.equals( replaceParam)) {
                // シート値
                paramValues = ReportsUtil.getSheetValues( reportsParserInfo.getReportBook(), propertyName, reportsParserInfo.getReportParsers()).toArray();
            } else {
                // 置換する値を取得
                ParamInfo paramInfo = reportsParserInfo.getParamInfo();
                if ( paramInfo != null) {
                    paramValues = getParamData( paramInfo, replaceParam);
                }
            }

            if ( paramValues == null || paramValues.length == 0) {
                // 空文字置換
                paramValues = new Object[] {null};
            }

            // シフトセル数を取得
            int shiftNum = paramValues.length;
            // パラメータ数を取得
            int paramLength = paramValues.length;

            // 対象セルの開始行番号を取得
            int defaultFromCellRowIndex = tagCell.getRowIndex();
            // 対象セルの開始列番号を取得
            int defaultFromCellColIndex = tagCell.getColumnIndex();

            // 単位セル行サイズ（初期値：１セル）
            int unitRowSize = 1;

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
                        // 単位セル行サイズに結合セル行数を設定する
                        unitRowSize = curMergedAdress.getLastRow() - curMergedAdress.getFirstRow() + 1;

                        // シフトセル数を単位セル行サイズ分（結合セル行数分）増加させる
                        shiftNum = shiftNum * unitRowSize;
                    }
                }
            }
            
            // ■シフト前のシート（フォーマット）情報を保存
            // 最終行座標までのセル情報を保存する
            tagCell = new CellClone( tagCell);
            List<Cell> cellList = new ArrayList<Cell>();
            int defaultToOverCellRowIndex = tagCell.getRowIndex() + unitRowSize;
            for ( int i = defaultFromCellRowIndex; i < defaultToOverCellRowIndex; i++) {
                Row targetCellRow = sheet.getRow(i);
                cellList.add( new CellClone( targetCellRow.getCell(tagCell.getColumnIndex())));
            }

            // ■シフトセル数の設定
            if ( repeatNum != null && repeatNum < shiftNum) {
                // 繰り返し最大回数が指定されており、
                // かつ、シフトセル数が繰り返し最大回数を越えてしまっていた場合には
                // 繰り返し最大回数×単位セル行サイズをシフトセル数とする
                shiftNum = repeatNum * unitRowSize;

                // パラメタ数を繰り返し最大回数に置き換える
                paramLength = repeatNum;
            }

            // ■シフト処理
            if ( shiftNum > 1) {
                // 値の数分シフト(１つであれば、シフトしない)
                int shiftRowSize = tagCell.getRowIndex() + shiftNum - unitRowSize - 1;
                if ( !rowShift) {
                    // セル単位でシフトする場合
                    CellRangeAddress rangeAddress = new CellRangeAddress( tagCell.getRowIndex(), shiftRowSize, tagCell.getColumnIndex(), tagCell.getColumnIndex());
                    PoiUtil.insertRangeDown( sheet, rangeAddress);
                } else {
                    // 行単位でシフトする場合
                    int shiftStartRow = tagCell.getRowIndex() + 1;
                    int shiftEndRow = sheet.getLastRowNum();
                    if ( shiftEndRow < shiftStartRow) {
                        // タグが終了行に存在していた場合には「最終行＝開始行＋１」に置き換える
                        // ※開始行＝終了行＋１となってしまうことを防ぐ
                        shiftEndRow = shiftStartRow + 1;
                    }
                    sheet.shiftRows( shiftStartRow, shiftEndRow, shiftNum - unitRowSize);
                }
            }

            // ■ブック、シートの取得
            Workbook workbook = sheet.getWorkbook();
            String sheetName = workbook.getSheetName( workbook.getSheetIndex( sheet));
            // シート名
            List<String> sheetNames = ReportsUtil.getSheetNames( reportsParserInfo.getReportBook());
            // 変換値
            List<Object> resultValues = new ArrayList<Object>();
            // 前回配列値(beforeValue)
            Object beforeValue = null;

            // 設定値配列インデックス
            int valueIndex = -1;
            // シフトセル数分まで行毎の設定処理を行う
            for ( int rowIndex = 0; rowIndex < shiftNum; rowIndex++) {
                // 行の取得
                Row row = sheet.getRow( tagCell.getRowIndex() + rowIndex);
                if ( row == null) {
                    row = sheet.createRow( tagCell.getRowIndex() + rowIndex);
                }
                // セルの取得
                Cell cell = row.getCell( tagCell.getColumnIndex());
                if ( cell == null) {
                    cell = row.createCell( tagCell.getColumnIndex());
                }
                // セルに設定する値の初期化(null)
                Object value = null;

                // セルの何番目に該当するかを除算にて算出
                // ※結果値が0の場合、(結合、非結合を問わず)最初のセルであることを示す
                int cellIndex = rowIndex % unitRowSize;
                
                // 結合行の値設定を行うセル以外は値を書き換える
                boolean skipRow = false;
                if ( cellIndex != 0) {
                    skipRow = true;
                } else {
                    valueIndex++;
                }

                // ■コピー(セル情報の設定)
                PoiUtil.copyCell( cellList.get(cellIndex), cell);

                // 値の設定
                Object currentValue = paramValues[valueIndex];
                // 重複非表示オプション=trueかつ、値が一致する場合はフラグを書き換える
                boolean duplicateValue = false;
                if ( beforeValue != null && currentValue != null && beforeValue.equals( currentValue)) {
                    // 値が一致する場合
                    duplicateValue = true;
                }
                if ( !skipRow && !(hideDuplicate && duplicateValue)) {
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
                    if ( !skipRow && valueIndex < sheetNames.size()) {
                        PoiUtil.setHyperlink( cell, Hyperlink.LINK_DOCUMENT, "'" + sheetNames.get( valueIndex) + "'!A1");
                        if ( log.isDebugEnabled()) {
                            log.debug( "[シート名=" + sheetName + ",セル=(" + cell.getRowIndex() + "," + cell.getColumnIndex() + ")]  Hyperlink ⇒ " + "'" + sheetNames.get( valueIndex) + "'!A1");
                        }
                    }
                }

                // ■改ページオプション指定
                if ( !skipRow) {
                    if ( breakNum != null && valueIndex != 0 && valueIndex % breakNum == 0) {
                        // データ数により改ページする場合
                        sheet.setRowBreak( row.getRowNum() - 1);
                    }
                    if ( changeBreak && valueIndex != 0 && !duplicateValue) {
                        // 値変更により改ページする場合
                        sheet.setRowBreak( row.getRowNum() - 1);
                    }
                }

                // ■結合セル指定
                // 元々テンプレートで結合済みのセルが存在するため、
                // 同じ結合セル情報は２重登録しないようにする。
                if ( !skipRow && unitRowSize > 1) {
                    // 結合フラグの設定
                    boolean mergedRegionFlag = false;
                    if ( rowShift && valueIndex != 0) {
                        // 行単位でシフトする場合、最初値は結合しない
                        mergedRegionFlag = true;
                    } else if ( !rowShift && paramLength > valueIndex + 1) {
                        // セル単位でシフトする場合、最終値は結合しない
                        mergedRegionFlag = true;
                    }

                    // 結合セル情報登録
                    if ( mergedRegionFlag) {
                        CellRangeAddress rangeAddress = new CellRangeAddress( cell.getRowIndex(), cell.getRowIndex() + unitRowSize - 1, cell.getColumnIndex(), cell.getColumnIndex());
                        sheet.addMergedRegion( rangeAddress);

                        // 結合の場合は結合したタイミングで前回の値を更新
                        beforeValue = currentValue;
                    }
                }

                // 結合では無い場合は毎回値を更新
                if ( unitRowSize == 1) {
                    beforeValue = currentValue;
                }

            }

            // 解析結果の生成
            ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
            parsedReportInfo.setParsedObject( resultValues);
            // 最終行を設定（行結合セル状態も考慮）
            parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex() + unitRowSize - 1);
            parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex());
            parsedReportInfo.setRowIndex( tagCell.getRowIndex() + shiftNum - 1);
            parsedReportInfo.setColumnIndex( tagCell.getColumnIndex());
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
     * @param paramDef パラメータ定義
     * @param tagCell タグセル
     * @throws ParseException 不正なパラメータがある場合
     */
    private void checkParam( Map<String, String> paramDef, Cell tagCell) throws ParseException {
        // シートハイパーリンク設定有無と重複非表示は重複不可
        if ( paramDef.containsKey( PARAM_DUPLICATE) && paramDef.containsKey( PARAM_SHEET_LINK)) {
            throw new ParseException( tagCell, "二重定義：" + PARAM_DUPLICATE + "," + PARAM_SHEET_LINK);
        }
    }

    /*
     * (non-Javadoc)
     * 
     * @see org.bbreak.excella.reports.tag.ReportsTagParser#useControlRow()
     */
    @Override
    public boolean useControlRow() {
        return false;
    }

}
