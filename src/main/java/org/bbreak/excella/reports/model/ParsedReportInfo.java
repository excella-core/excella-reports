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


/**
 * 帳票情報の解析結果を保持するクラス
 * 
 * @since 1.0
 */
@SuppressWarnings("serial")
public class ParsedReportInfo implements Serializable{

    /**
     * 処理最終行番号
     */
    private int rowIndex = -1;

    /**
     * 処理最終列番号
     */
    private int columnIndex = -1;

    /**
     * 処理前最終行番号
     */
    private int defaultRowIndex = -1;
    
    /**
     * 処理前最終列番号
     */
    private int defaultColumnIndex = -1;

    /**
     * 解析結果オブジェクト
     */
    private Object parsedObject = null;
    
    
    /**
     * デフォルトコンストラクタ
     */
    public ParsedReportInfo(){
    }
    

    /**
     * 処理最終行番号を取得します。
     * 
     * @return 処理最終行番号
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * 処理最終行番号を設定します。
     * 
     * @param rowIndex 処理最終行番号
     */
    public void setRowIndex( int rowIndex) {
        this.rowIndex = rowIndex;
    }

    /**
     * 処理最終列番号を取得します。
     * 
     * @return 処理最終列番号
     */
    public int getColumnIndex() {
        return columnIndex;
    }

    /**
     * 処理最終列番号を設定します。
     * 
     * @param columnIndex 処理最終列番号
     */
    public void setColumnIndex( int columnIndex) {
        this.columnIndex = columnIndex;
    }

    /**
     * 解析結果オブジェクトを取得します。
     * 
     * @return 解析結果オブジェクト
     */
    public Object getParsedObject() {
        return parsedObject;
    }

    /**
     * 解析結果オブジェクトを設定します。
     * 
     * @param parsedObject 解析結果オブジェクト
     */
    public void setParsedObject( Object parsedObject) {
        this.parsedObject = parsedObject;
    }

    /*
     * (non-Javadoc)
     * 
     * @see java.lang.Object#toString()
     */
    @Override
    public String toString() {

        StringBuilder sb = new StringBuilder();
        sb.append( "beforeLastCell=(");
        sb.append( defaultRowIndex).append( ",").append( defaultColumnIndex).append( "), ");
        sb.append( "afterLastCell=(");
        sb.append( rowIndex).append( ",").append( columnIndex).append( "), ");
        sb.append( parsedObject);

        return sb.toString();
    }

    /**
     * 処理前最終行番号を取得します。
     * 
     * @return 処理前最終行番号
     */
    public int getDefaultRowIndex() {
        return defaultRowIndex;
    }

    /**
     * 処理前最終行番号を設定します。
     * 
     * @param defaultRowIndex 処理前最終行番号
     */
    public void setDefaultRowIndex( int defaultRowIndex) {
        this.defaultRowIndex = defaultRowIndex;
    }

    /**
     * 処理前最終列番号を取得します。
     * 
     * @return 処理前最終列番号
     */
    public int getDefaultColumnIndex() {
        return defaultColumnIndex;
    }

    /**
     * 処理前最終列番号を設定します。
     * 
     * @param defaultColumnIndex 処理前最終列番号
     */
    public void setDefaultColumnIndex( int defaultColumnIndex) {
        this.defaultColumnIndex = defaultColumnIndex;
    }

}
