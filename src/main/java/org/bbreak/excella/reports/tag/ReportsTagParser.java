/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportsTagParser.java 165 2010-08-05 11:01:01Z ogiharasf $
 * $Revision: 165 $
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
