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
