/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: RemoveAdapter.java 199 2012-12-17 14:47:10Z akira-yokoi $
 * $Revision: 199 $
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
package org.bbreak.excella.reports.listener;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.tag.RemoveParamParser;
import org.bbreak.excella.reports.tag.ReportsTagParser;

/**
 * シート解析後にセル・列・行・制御行を削除するアダプタ
 * 
 * @since 1.0
 */
public class RemoveAdapter extends ReportProcessAdaptor {

    /**
     * 行
     */
    public static final String ROW = "row";

    /**
     * 列
     */
    public static final String COLUMN = "column";

    /**
     * セル
     */
    public static final String CELL = "cell";

    /**
     * 左方向にシフト（デフォルト）
     */
    public static final String LEFT = "left";

    /**
     * 上方向にシフト
     */
    public static final String UP = "up";

    /*
     * (non-Javadoc)
     * 
     * @see org.bbreak.excella.reports.listener.ReportProcessAdaptor#postParse(org.apache.poi.ss.usermodel.Sheet, org.bbreak.excella.core.SheetParser, org.bbreak.excella.core.SheetData)
     */
    @Override
    public void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException {

        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        for ( int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {

            Row row = sheet.getRow( rowIndex);
            if ( row != null) {
                int firstColNum = row.getFirstCellNum();
                int lastColNum = row.getLastCellNum() - 1;
                boolean isRowFlag = false;

                for ( int colIndex = firstColNum; colIndex <= lastColNum; colIndex++) {
                    Cell cell = row.getCell( colIndex);
                    if ( cell != null) {
                        if ( cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().contains( RemoveParamParser.DEFAULT_TAG)) {
                            // タグのパラメータを取得
                            String[] paramArray = getStrParam( sheet, rowIndex, colIndex);

                            // 削除単位（セル・列・行）
                            String removeUnit = paramArray[0];
                            // タグを持つセルを削除
                            row.removeCell( cell);

                            // 行全体削除の場合の処理
                            if ( removeUnit.equals( "") || removeUnit.equals( ROW)) {
                                removeRegion( sheet, rowIndex, -1);
                                removeControlRow( sheet, rowIndex);
                                isRowFlag = true;
                                break;
                            } else if ( removeUnit.equals( CELL) || removeUnit.equals( COLUMN)) {
                                // セルまたは列全体削除の場合の処理
                                removeCellOrCol( paramArray, removeUnit, sheet, row, cell, rowIndex, colIndex);
                            }
                            lastColNum = row.getLastCellNum() - 1;
                            colIndex--;
                        }
                        // 制御行の処理
                        if ( isControlRow( sheet, sheetParser, row, cell)) {
                            removeControlRow( sheet, rowIndex);
                            isRowFlag = true;
                            break;
                        }
                    }
                }
                // 行を削除した場合
                if ( isRowFlag) {
                    lastRowNum = sheet.getLastRowNum();
                    rowIndex--;
                }
            }
        }
    }

    /**
     * REMOVEタグのパラメータを取得する
     * 
     * @param sheet
     * @param rowIndex
     * @param colIndex
     * @return paramArray
     */
    private String[] getStrParam( Sheet sheet, int rowIndex, int colIndex) throws ParseException {
        Row row = sheet.getRow( rowIndex);
        Cell cell = row.getCell( colIndex);
        String strValue = cell.getStringCellValue();

        // タグのパラメータ取得
        String param = TagUtil.getParam( strValue);
        String[] paramArray = param.split( TagParser.PARAM_DELIM);

        return paramArray;
    }

    /**
     * セルの削除、または列の削除を行う
     * 
     * @param paramArray
     * @param removeUnit
     * @param sheet
     * @param row
     * @param cell
     * @param rowIndex
     * @param colIndex
     * @return paramArray
     */
    private void removeCellOrCol( String[] paramArray, String removeUnit, Sheet sheet, Row row, Cell cell, int rowIndex, int colIndex) {
        if ( removeUnit.equals( CELL)) {
            removeRegion( sheet, rowIndex, colIndex);
            if ( paramArray.length > 1) {
                Row removeRow = sheet.getRow( rowIndex);
                if( removeRow != null){
                    Cell removeCell = removeRow.getCell( colIndex);
                    if( removeCell != null){
                        removeRow.removeCell( removeCell);                        
                    }
                }
                
                // セル削除後のシフト方向
                String direction = paramArray[1];
                if ( direction.equals( LEFT)) {
                    shiftLeft( row, cell, colIndex);
                } else if ( direction.equals( UP)) {
                    shiftUp( sheet, cell, rowIndex, colIndex);
                }
            } else {
                // 移動方向の指定がない時は左方向にシフトする
                shiftLeft( row, cell, colIndex);
            }
        } else if ( removeUnit.equals( COLUMN)) {
            removeRegion( sheet, -1, colIndex);
            // 列全体削除時の処理
            for ( int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row removeCol = sheet.getRow( rowNum);
                if ( removeCol != null) {
                    Cell removeCell = removeCol.getCell( colIndex);
                    if ( removeCell != null) {
                        removeCol.removeCell( removeCell);
                    }
                }
                shiftLeft( sheet.getRow( rowNum), cell, colIndex);
            }
        }
    }

    /**
     * セル削除に伴いセルを左方向に動かす（デフォルト）
     * 
     * @param row
     * @param cell
     * @param colIndex
     */
    private void shiftLeft( Row row, Cell cell, int colIndex) {
        // コピー・貼付
        int startCopyIndex = colIndex + 1;
        if ( row == null) {
            return;
        }
        int finishCopyIndex = row.getLastCellNum() - 1;

        for ( int copyColNum = startCopyIndex; copyColNum <= finishCopyIndex; copyColNum++) {
            // コピー元セル
            Cell fromCell = row.getCell( copyColNum);
            // コピー先セル
            Cell toCell = row.getCell( copyColNum - 1);

            if ( fromCell != null) {
                if ( toCell == null) {
                    toCell = row.createCell( copyColNum - 1);
                }
                PoiUtil.copyCell( fromCell, toCell);
                row.removeCell( fromCell);
            }
        }
    }

    /**
     * セル削除に伴いセルを上方向に動かす
     * 
     * @param sheet
     * @param cell
     * @param rowIndex
     * @param colIndex
     */
    private void shiftUp( Sheet sheet, Cell cell, int rowIndex, int colIndex) {
        // コピー開始・終了行番号
        int startCopyIndex = rowIndex + 1;
        int finishCopyIndex = sheet.getLastRowNum();

        for ( int copyRowNum = startCopyIndex; copyRowNum <= finishCopyIndex; copyRowNum++) {

            Row row = sheet.getRow( copyRowNum);
            if ( row != null) {
                Row preRow = sheet.getRow( copyRowNum - 1);
                // コピー元セル
                Cell fromCell = row.getCell( colIndex);

                if ( fromCell != null) {
                    // コピー先セル
                    Cell toCell = null;
                    if ( preRow == null) {
                        preRow = sheet.createRow( copyRowNum - 1);
                    }
                    toCell = preRow.getCell( colIndex);
                    if ( toCell == null) {
                        toCell = preRow.createCell( colIndex);
                    }
                    PoiUtil.copyCell( fromCell, toCell);
                    row.removeCell( fromCell);
                }
            }
        }
    }

    /**
     * 行を削除する
     * 
     * @param sheet
     * @param rowIndex
     */
    private void removeControlRow( Sheet sheet, int rowIndex) {

        // 最終行だったら、削除
        if ( rowIndex == sheet.getLastRowNum()) {
            sheet.removeRow( sheet.getRow( rowIndex));
        } else {
            sheet.removeRow( sheet.getRow( rowIndex));
            sheet.shiftRows( rowIndex + 1, sheet.getLastRowNum(), -1, true, true);
        }
    }

    /**
     * 対象の行、列番号が結合されている場合は結合を解除する
     * 
     * @param sheet
     * @param removeRowNum
     * @param removeColNum
     */
    private void removeRegion( Sheet sheet, int removeRowNum, int removeColNum) {
        // シート内の結合情報の数だけループする
        for ( int i = 0; i < sheet.getNumMergedRegions(); i++) {
            // 結合情報(Region)を取得する
            CellRangeAddress region = sheet.getMergedRegion( i);

            if ( removeRowNum != -1) {
                if ( region.getFirstRow() > removeRowNum || removeRowNum > region.getLastRow()) {
                    continue;
                }
            }

            if ( removeColNum != -1) {
                if ( region.getFirstColumn() > removeColNum || removeColNum > region.getLastColumn()) {
                    continue;
                }
            }
            // 引数の結合INDEXに該当する結合を解除する
            sheet.removeMergedRegion( i);
            return;
        }
    }

    /**
     * 制御行の有無を返す。
     * 
     * @param sheet
     * @param sheetParser
     * @param row
     * @return
     */
    private boolean isControlRow( Sheet sheet, SheetParser sheetParser, Row row, Cell cell) {

        List<TagParser<?>> parsers = sheetParser.getTagParsers();

        for ( TagParser<?> parser : parsers) {
            try {
                if ( parser.isParse( sheet, cell)) {
                    if ( parser instanceof ReportsTagParser) {
                        ReportsTagParser<?> reportsTagParser = ( ReportsTagParser<?>) parser;
                        if ( reportsTagParser.useControlRow()) {
                            return true;
                        }
                    }
                }
            } catch ( ParseException e) {
                // 解析できないものは扱わない
                continue;
            }
        }
        return false;
    }
}
