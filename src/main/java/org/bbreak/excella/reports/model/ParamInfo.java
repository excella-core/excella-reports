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
