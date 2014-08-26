/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: SumParamParser.java 5 2009-06-22 07:55:44Z tomo-shibata $
 * $Revision: 5 $
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

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.ReportsParserInfo;
import org.bbreak.excella.reports.util.ReportsUtil;

/**
 * シート内の置換文字列を変換するパーサ
 * 
 * @since 1.0
 */
public class SumParamParser extends ReportsTagParser<Object> {

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$SUM";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * コンストラクタ
     */
    public SumParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public SumParamParser( String tag) {
        super( tag);
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {

        Map<String, String> paramDef = TagUtil.getParams( tagCell.getStringCellValue());

        ReportsParserInfo reportsParserInfo = ( ReportsParserInfo) data;

        ParamInfo paramInfo = reportsParserInfo.getParamInfo();

        Object paramValue = null;

        if ( paramInfo != null) {
            // 置換
            // 置換変数の取得
            String replaceParam = paramDef.get( PARAM_VALUE);

            paramValue = ReportsUtil.getSumValue( paramInfo, replaceParam, reportsParserInfo.getReportParsers());

            // 置換
            PoiUtil.setCellValue( tagCell, paramValue);

        }

        // 解析結果の生成
        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
        parsedReportInfo.setParsedObject( paramValue);
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
