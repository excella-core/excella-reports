/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ParamInfo.java 194 2010-08-31 09:34:14Z akira-yokoi $
 * $Revision: 194 $
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
package org.bbreak.excella.reports.model;

import java.io.Serializable;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.Map.Entry;

/**
 * テンプレートを置換するパラメータ情報を保持するクラス
 * 
 * @since 1.0
 */
@SuppressWarnings( "serial")
public class ParamInfo implements Serializable {

    /**
     * キーのデリミタ
     */
    private static final String KEY_DELIMITER = ":";

    /**
     * デフォルトコンストラクタ
     */
    public ParamInfo() {
    }

    /**
     * パラメータデータのマップ<BR>
     * キー：タグ + {@link #KEY_DELIMITER} + パラメータ名
     */
    private Map<String, Object> paramDataMap = new HashMap<String, Object>();

    /**
     * パラメータのデータ取得
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     * @return パラメータ名に対応するデータ。存在しない場合はnull
     */
    public Object getParam( String tag, String paramName) {
        return paramDataMap.get( tag + KEY_DELIMITER + paramName);
    }

    /**
     * パラメータの追加
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     * @param data データ
     */
    public void addParam( String tag, String paramName, Object data) {
        paramDataMap.put( tag + KEY_DELIMITER + paramName, data);
    }

    /**
     * パラメータの追加
     * 
     * @param tag タグ名
     * @param params 追加パラメータ
     */
    public void addParams( String tag, Map<String, Object> params) {
        for ( Entry<String, Object> entry : params.entrySet()) {
            addParam( tag, entry.getKey(), entry.getValue());
        }
    }

    /**
     * パラメータの削除
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     */
    public void removeParam( String tag, String paramName) {
        paramDataMap.remove( tag + KEY_DELIMITER + paramName);
    }

    /**
     * パラメータのクリア
     * 
     * @param tag タグ名
     */
    public void clearParam( String tag) {
        Set<String> removeKeys = new HashSet<String>();

        for ( String key : paramDataMap.keySet()) {
            if ( key.startsWith( tag + KEY_DELIMITER)) {
                removeKeys.add( key);
            }
        }
        for ( String key : removeKeys) {
            paramDataMap.remove( key);
        }

    }

    /**
     * パラメータのクリア
     */
    public void clearParam() {
        paramDataMap.clear();
    }

    /**
     * パラメータの取得
     * 
     * @return パラメータのマップ
     */
    public Map<String, Object> getParams() {
        return paramDataMap;
    }

    /**
     * パラメータの設定
     * 
     * @param params 設定するパラメータ
     */
    public void setParams( Map<String, Object> params) {
        paramDataMap.putAll( params);
    }
}
