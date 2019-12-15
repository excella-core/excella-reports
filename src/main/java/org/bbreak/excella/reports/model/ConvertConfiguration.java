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
import java.util.Map;
import java.util.Set;

/**
 * 変換情報を保持するクラス
 * 
 * @since 1.0
 */
@SuppressWarnings("serial")
public class ConvertConfiguration implements Serializable{

    /**
     * フォーマット
     */
    private String formatType = null;

    /**
     * 出力オプション
     */
    private Map<String, Object> options = new HashMap<String, Object>();
    
    /**
     * デフォルトコンストラクタ
     */
    public ConvertConfiguration(){
    }
    
    /**
     * コンストラクタ
     * 
     * @param formatType フォーマット
     */
    public ConvertConfiguration( String formatType) {
        this.formatType = formatType;
    }

    /**
     * オプションを追加する。
     * 
     * @param property プロパティ
     * @param value 値
     */
    public void addOption( String property, Object value) {
        options.put( property, value);
    }

    /**
     * オプションを追加する。
     * 
     * @param options オプション
     */
    public void addOptions( Map<String, Object> options) {
        this.options.putAll( options);
    }

    /**
     * オプションを初期化する。
     */
    public void clearOptions() {
        options.clear();
    }

    /**
     * オプションを削除する。
     * 
     * @param property プロパティ
     */
    public void removeOption( String property) {
        options.remove( property);
    }

    /**
     * オプションプロパティ一覧の取得する。
     * 
     * @return プロパティ一覧
     */
    public Set<String> getOptionsProperties() {
        return options.keySet();
    }

    /**
     * オプション設定値を取得する。
     * 
     * @param property プロパティ
     * @return オプション設定値
     */
    public Object getOptionValue( String property) {
        return options.get( property);
    }

    /**
     * オプションを取得する。
     * 
     * @return オプション
     */
    public Map<String, Object> getOptions() {
        return new HashMap<String, Object>( options);
    }

    /**
     * フォーマットを取得します。
     * 
     * @return フォーマット
     */
    public String getFormatType() {
        return formatType;
    }

    /**
     * フォーマットを設定します。
     * 
     * @param formatType フォーマット
     */
    public void setFormatType( String formatType) {
        this.formatType = formatType;
    }

}
