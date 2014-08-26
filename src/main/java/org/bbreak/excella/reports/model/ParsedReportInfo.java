/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ParsedReportInfo.java 187 2010-08-17 17:02:35Z ogiharasf $
 * $Revision: 187 $
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
