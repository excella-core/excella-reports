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
package org.bbreak.excella.reports.tag;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.model.ParsedReportInfo;

/**
 * 改ページするパーサー
 * 
 * @author T.Maruyama
 */
public class BreakParamParser extends ReportsTagParser<Object> {

    /** デフォルトタグ */
    public static final String DEFAULT_TAG = "$BREAK";

    /**
     * コンストラクタ
     */
    public BreakParamParser() {
        this( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag
     */
    public BreakParamParser( String tag) {
        super( tag);
    }

    /**
     * @see org.bbreak.excella.reports.tag.ReportsTagParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)
     */
    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
        // このタグは何も出力しない
        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
        parsedReportInfo.setRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setColumnIndex( tagCell.getColumnIndex());
        parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex());
        return parsedReportInfo;
    }

    /**
     * @see org.bbreak.excella.reports.tag.ReportsTagParser#useControlRow()
     */
    @Override
    public boolean useControlRow() {
        return false;
    }

}
