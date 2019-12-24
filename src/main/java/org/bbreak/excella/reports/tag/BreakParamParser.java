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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.model.ParsedReportInfo;

/**
 * 改ページするパーサー
 * 
 * @author T.Maruyama
 * @version 1.8
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
