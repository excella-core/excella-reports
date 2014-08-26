/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ConvertConfiguration.java 137 2010-07-29 04:35:30Z akira-yokoi $
 * $Revision: 137 $
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
