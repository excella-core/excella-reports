/*************************************************************************
 *
 * Copyright 2015 by bBreak Systems.
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
 * version 3 along with ExCella Reports.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.reports.listener;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.tag.BreakParamParser;

/**
 * 改ページの挿入
 * 
 * @author T.Maruyama
 * @version 1.8
 */
public class BreakAdapter extends ReportProcessAdaptor {

    /**
     * @see org.bbreak.excella.reports.listener.ReportProcessAdaptor#postParse(org.apache.poi.ss.usermodel.Sheet, org.bbreak.excella.core.SheetParser, org.bbreak.excella.core.SheetData)
     */
    @Override
    public void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException {
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();

        for ( int rowIndex = firstRowNum; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow( rowIndex);
            if ( row != null) {
                parseRow( sheet, sheetParser, sheetData, row, rowIndex);
            }
        }
    }

    /**
     * 行単位で解析し、必要なら改ページを挿入する
     */
    protected void parseRow( Sheet sheet, SheetParser sheetParser, SheetData sheetData, Row row, int rowIndex) {
        int firstColNum = row.getFirstCellNum();
        int lastColNum = row.getLastCellNum() - 1;

        for ( int colIndex = firstColNum; colIndex <= lastColNum; colIndex++) {
            Cell cell = row.getCell( colIndex);
            if ( cell != null) {
                if ( cell.getCellTypeEnum() == CellType.STRING && cell.getStringCellValue().contains( BreakParamParser.DEFAULT_TAG)) {
                    // 改ページを挿入
                    if ( isInMergedRegion( sheet, row, cell)) {
                        setRowBreakMergedRegion( sheet, row, cell);
                    } else {
                        setRowBreak( sheet, row, cell);
                    }
                }
            }
        }

    }

    private boolean isInMergedRegion( Sheet sheet, Row row, Cell cell) {
        // シート内の結合情報の数だけループする
        for ( int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion( i);
            if ( region.isInRange( row.getRowNum(), cell.getColumnIndex())) {
                return true;
            }
        }
        return false;
    }

    protected void setRowBreakMergedRegion( Sheet sheet, Row row, Cell cell) {
        PoiUtil.setCellValue( cell, null);
        // シート内の結合情報の数だけループする
        for ( int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion( i);
            if ( region.isInRange( row.getRowNum(), cell.getColumnIndex())) {
                // 結合されていれば結合セルの直後に改ページ
                sheet.setRowBreak( row.getRowNum() + 1);
                return;
            }
        }
        // 結合されていなければその行で改ページ
        sheet.setRowBreak( row.getRowNum());
    }

    protected void setRowBreak( Sheet sheet, Row row, Cell cell) {
        sheet.setRowBreak( row.getRowNum());
        row.removeCell( cell);
    }
}
