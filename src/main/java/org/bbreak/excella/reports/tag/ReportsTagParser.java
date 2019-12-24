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
import org.bbreak.excella.core.tag.TagParser;
import org.bbreak.excella.reports.model.ParamInfo;
import org.bbreak.excella.reports.model.ParsedReportInfo;

/**
 * 帳票処理のパーサークラス
 * 
 * @param <T> パラメータを置換する値の型
 * @since 1.0
 */
public abstract class ReportsTagParser<T> extends TagParser<ParsedReportInfo> {

    /*
     * (non-Javadoc)
     * 
     * @see org.bbreak.excella.core.tag.TagParser#parse(org.apache.poi.ss.usermodel.Sheet, org.apache.poi.ss.usermodel.Cell, java.lang.Object)
     */
    @Override
    public abstract ParsedReportInfo parse( Sheet sheet, Cell tagCell, Object data) throws ParseException;

    /**
     * コンストラクタ
     * 
     * @param tag タグ
     */
    public ReportsTagParser( String tag) {
        super( tag);
    }

    /**
     * タグを制御行として扱うか否かを取得する。
     * 
     * @return true:制御行として削除／false:置換
     */
    public abstract boolean useControlRow();

    /**
     * パラメータ情報よりパラメータ名で格納されている値を取得する。
     * 
     * @param paramInfo パラメータ情報
     * @param paramName パラメータ名
     * @return 置換する値
     */
    @SuppressWarnings( "unchecked")
    public T getParamData( ParamInfo paramInfo, String paramName) {
        if ( paramInfo == null || paramName == null) {
            return null;
        }
        return ( T) paramInfo.getParam( getTag(), paramName);
    }

}
