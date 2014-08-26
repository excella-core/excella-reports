/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: RemoveParamParser.java 57 2009-10-19 11:41:49Z a-hoshino $
 * $Revision: 57 $
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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.model.ParsedReportInfo;

/**
 * セル・列・行を削除するパーサー
 * 
 * @since 1.0
 */
public class RemoveParamParser extends ReportsTagParser<Object> {

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$REMOVE";

    /**
     * コンストラクタ
     */
    public RemoveParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public RemoveParamParser( String tag) {
        super( tag);
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
        // 解析結果の生成
        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
        parsedReportInfo.setRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setColumnIndex( tagCell.getColumnIndex());
        parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex());
        return parsedReportInfo;
    }

    @Override
    public boolean useControlRow() {
        return false;
    }

}
