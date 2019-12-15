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
