/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportSheet.java 136 2010-07-29 04:29:25Z akira-yokoi $
 * $Revision: 136 $
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
import java.util.Map;

/**
 * ワークシートの置換情報を保持するクラス
 * 
 * @since 1.0
 */
@SuppressWarnings("serial")
public class ReportSheet implements Serializable{

    /**
     * テンプレートシート名
     */
    private String templateName = null;

    /**
     * 出力シート名
     */
    private String sheetName = null;

    /**
     * パラメータ情報
     */
    private ParamInfo paramInfo = new ParamInfo();

    /**
     * デフォルトコンストラクタ
     */
    public ReportSheet() {
    }

    /**
     * コンストラクタ<BR>
     * テンプレート名≠出力シート名の場合
     * @param templateName テンプレートシート名
     * @param sheetName 出力シート名
     */
    public ReportSheet( String templateName, String sheetName) {
        this.templateName = templateName;
        this.sheetName = sheetName;
    }
    
    /**
     * コンストラクタ<BR>
     * テンプレート名＝シート名の場合
     * @param sheetName シート名
     */
    public ReportSheet( String sheetName) {
        this.templateName = sheetName;
        this.sheetName = sheetName;
    }

    /**
     * テンプレートシート名を取得します。
     * 
     * @return テンプレートシート名
     */
    public String getTemplateName() {
        return templateName;
    }

    /**
     * 出力シート名を取得します。
     * 
     * @return 出力シート名
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * 出力シート名を設定します。
     * 
     * @param sheetName 出力シート名
     */
    public void setSheetName( String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * テンプレートシート名を設定します。
     * 
     * @param templateName テンプレートシート名
     */
    public void setTemplateName( String templateName) {
        this.templateName = templateName;
    }

    /**
     * 置換パラメータ情報を取得する
     * @return 置換パラメータ情報　
     */
    public ParamInfo getParamInfo() {
        return paramInfo;
    }

    /**
     * パラメータの追加
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     * @param data データ
     * @see org.bbreak.excella.reports.model.ParamInfo#addParam(java.lang.String, java.lang.String, java.lang.Object)
     */
    public void addParam( String tag, String paramName, Object data) {
        paramInfo.addParam( tag, paramName, data);
    }

    /**
     * パラメータの追加
     * 
     * @param tag タグ名
     * @param params 追加パラメータ
     * @see org.bbreak.excella.reports.model.ParamInfo#addParams(java.lang.String, java.util.Map)
     */
    public void addParams( String tag, Map<String, Object> params) {
        paramInfo.addParams( tag, params);
    }

    /**
     * パラメータのクリア
     * @see org.bbreak.excella.reports.model.ParamInfo#clearParam()
     */
    public void clearParam() {
        paramInfo.clearParam();
    }

    /**
     * パラメータのクリア
     * 
     * @param tag タグ名
     * @see org.bbreak.excella.reports.model.ParamInfo#clearParam(java.lang.String)
     */
    public void clearParam( String tag) {
        paramInfo.clearParam( tag);
    }

    /**
     * パラメータのデータ取得
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     * @return パラメータ名に対応するデータ。存在しない場合はnull
     * @see org.bbreak.excella.reports.model.ParamInfo#getParam(java.lang.String, java.lang.String)
     */
    public Object getParam( String tag, String paramName) {
        return paramInfo.getParam( tag, paramName);
    }

    /**
     * パラメータの削除
     * 
     * @param tag タグ名
     * @param paramName パラメータ名
     * @see org.bbreak.excella.reports.model.ParamInfo#removeParam(java.lang.String, java.lang.String)
     */
    public void removeParam( String tag, String paramName) {
        paramInfo.removeParam( tag, paramName);
    }

}
