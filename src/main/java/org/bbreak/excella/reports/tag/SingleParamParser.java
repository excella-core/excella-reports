/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: SingleParamParser.java 70 2009-10-27 04:57:02Z akira-yokoi $
 * $Revision: 70 $
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

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.StringTokenizer;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.core.util.TagUtil;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;
import org.bbreak.excella.reports.processor.ReportsParserInfo;

/**
 * シート内の置換文字列を変換するパーサ
 * 
 * @since 1.0
 */
public class SingleParamParser extends ReportsTagParser<Object> {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( SingleParamParser.class);

    /**
     * デフォルトタグ
     */
    public static final String DEFAULT_TAG = "$";

    /**
     * 置換変数のパラメータ
     */
    public static final String PARAM_VALUE = "";

    /**
     * コンストラクタ
     */
    public SingleParamParser() {
        super( DEFAULT_TAG);
    }

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public SingleParamParser( String tag) {
        super( tag);
    }

    @Override
    public ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException {
        List<String> tmpTargets = new ArrayList<String>();
        StringTokenizer tagTokenizer = new StringTokenizer( tagCell.getStringCellValue(), getTag());
        while ( tagTokenizer.hasMoreTokens()) {
            tmpTargets.add( tagTokenizer.nextToken());
        }

        List<String> finalTargets = new ArrayList<String>();
        for ( String tempTarget : tmpTargets) {
            if ( tempTarget.startsWith( TagParser.TAG_PARAM_PREFIX) && !tempTarget.endsWith( TagParser.TAG_PARAM_SUFFIX)) {
                finalTargets.add( tempTarget.substring( 0, tempTarget.indexOf( TagParser.TAG_PARAM_SUFFIX) + 1));
                finalTargets.add( tempTarget.substring( tempTarget.indexOf( TagParser.TAG_PARAM_SUFFIX) + 1));
                continue;
            }
            finalTargets.add( tempTarget);
        }

        ParsedReportInfo parsedReportInfo = new ParsedReportInfo();
        List<Object> paramValues = new ArrayList<Object>();

        for ( int i = 0; i < finalTargets.size(); i++) {
            String target = finalTargets.get( i);
            if ( !target.startsWith( TagParser.TAG_PARAM_PREFIX)) {
                // 置換不要文字列
                paramValues.add( target);
                continue;
            } else {
                // 置換必要のためタグを追加
                target = getTag() + target;
            }

            Map<String, String> paramDef = TagUtil.getParams( target);

            // パラメータチェック
            checkParam( paramDef, tagCell);

            ReportsParserInfo reportsParserInfo = ( ReportsParserInfo) data;

            ParamInfo paramInfo = reportsParserInfo.getParamInfo();

            if ( paramInfo != null) {
                // 置換変数の取得
                String replaceParam = paramDef.get( PARAM_VALUE);

                // 置換する値を取得
                paramValues.add( getParamData( paramInfo, replaceParam));

                if ( log.isDebugEnabled()) {
                    Workbook workbook = sheet.getWorkbook();
                    String sheetName = workbook.getSheetName( workbook.getSheetIndex( sheet));
                    log.debug( "[シート名=" + sheetName + ",セル=(" + tagCell.getRowIndex() + "," + tagCell.getColumnIndex() + ")] " + tagCell.getStringCellValue() + " ⇒ " + paramValues.get( i));
                }
            }
        }

        // 置換
        if ( paramValues.size() > 1) {
            StringBuilder builder = new StringBuilder();
            for ( Object object : paramValues) {
                if ( object == null) {
                    continue;
                }
                if ( object instanceof Date) {
                    SimpleDateFormat sdf = new SimpleDateFormat( "yyyy/MM/dd");
                    builder.append( sdf.format( object));
                    continue;
                }
                builder.append( object);
            }
            PoiUtil.setCellValue( tagCell, builder.toString());
            parsedReportInfo.setParsedObject( builder.toString());
        } else if ( paramValues.size() == 1) {
            PoiUtil.setCellValue( tagCell, paramValues.get( 0));
            parsedReportInfo.setParsedObject( paramValues.get( 0));
        }
        parsedReportInfo.setDefaultRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setDefaultColumnIndex( tagCell.getColumnIndex());
        parsedReportInfo.setRowIndex( tagCell.getRowIndex());
        parsedReportInfo.setColumnIndex( tagCell.getColumnIndex());

        return parsedReportInfo;
    }

    /**
     * 不正なパラメータがある場合、ParseExceptionをthrowする。
     * 
     * @param paramDef パラメータ定義
     * @param tagCell タグセル
     * @throws ParseException 不正なパラメータがある場合
     */
    private void checkParam( Map<String, String> paramDef, Cell tagCell) throws ParseException {
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

    @Override
    public boolean isParse( Sheet sheet, Cell tagCell) throws ParseException {
        // 文字列かつ、タグを含むセルの場合は処理対象
        if ( tagCell.getCellType() == Cell.CELL_TYPE_STRING) {
            String cellTag = TagUtil.getTag( tagCell.getStringCellValue());
            if ( cellTag.endsWith( getTag())) {
                return true;
            }
        }
        return false;
    }
}
