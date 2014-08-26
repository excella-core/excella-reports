/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportsParserInfo.java 5 2009-06-22 07:55:44Z tomo-shibata $
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
package org.bbreak.excella.reports.processor;

import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.tag.ReportsTagParser;

/**
 * 帳票解析処理情報を保持するクラス
 * 
 * @since 1.0
 */
public class ReportsParserInfo {
    
    /**
     * ワークブックの置換情報
     */
    private ReportBook reportBook = null;

    /**
     * パラメータ情報
     */
    private ParamInfo paramInfo = null;

    /**
     * 使用パーサー
     */
    private List<ReportsTagParser<?>> reportParsers = null;

    /**
     * パラメータ情報を取得します。
     * @return パラメータ情報
     */
    public ParamInfo getParamInfo() {
        return paramInfo;
    }

    /**
     * パラメータ情報を設定します。
     * @param paramInfo パラメータ情報
     */
    public void setParamInfo( ParamInfo paramInfo) {
        this.paramInfo = paramInfo;
    }

    /**
     * 使用パーサーを取得します。
     * @return 使用パーサー
     */
    public List<ReportsTagParser<?>> getReportParsers() {
        return reportParsers;
    }

    /**
     * 使用パーサーを設定します。
     * @param reportParsers 使用パーサー
     */
    public void setReportParsers( List<ReportsTagParser<?>> reportParsers) {
        this.reportParsers = reportParsers;
    }

    /**
     * システム情報と帳票パーサーが同一の、帳票解析情報を作成する。
     * 
     * @param paramInfo 置換パラメータ情報
     * @return 帳票解析情報
     */
    public ReportsParserInfo createChildParserInfo( ParamInfo paramInfo) {
        ReportsParserInfo reportsParserInfo = new ReportsParserInfo();
        reportsParserInfo.reportParsers = reportParsers;
        reportsParserInfo.reportBook = reportBook;
        reportsParserInfo.paramInfo = paramInfo;

        return reportsParserInfo;

    }

    /**
     * 解析処理を行うパーサーを取得する。
     * 
     * @param sheet 対象シート
     * @param tagCell 対象セル
     * @return 解析処理を行うパーサー
     * @throws ParseException
     */
    public TagParser<?> getMatchTagParser( Sheet sheet, Cell tagCell) throws ParseException {
        if ( reportParsers != null) {
            for ( TagParser<?> tagParser : reportParsers) {
                if ( tagParser.isParse( sheet, tagCell)) {
                    
//                    TagParser<?> newParser;
//                    try {
//                        Constructor<? extends TagParser> cunstructor = tagParser.getClass().getConstructor( new Class[] {String.class});
//                        newParser = ( TagParser<?>) cunstructor.newInstance( tagParser.getTag());
//                    } catch ( Exception e) {
//                        throw new ParseException( tagCell, e);
//                    }
//
//                    return newParser;
                    return tagParser;
                }
            }
        }
        return null;
    }

    /**
     * ワークブックの置換情報を取得します。
     * @return ワークブックの置換情報
     */
    public ReportBook getReportBook() {
        return reportBook;
    }

    /**
     * ワークブックの置換情報を設定します。
     * @param reportBook ワークブックの置換情報
     */
    public void setReportBook( ReportBook reportBook) {
        this.reportBook = reportBook;
    }

}
