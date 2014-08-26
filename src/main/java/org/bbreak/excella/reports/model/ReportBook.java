/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportBook.java 136 2010-07-29 04:29:25Z akira-yokoi $
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
import java.net.URL;
import java.util.ArrayList;
import java.util.List;

/**
 * ワークブックの置換情報を保持するクラス
 * 
 * @since 1.0
 */
@SuppressWarnings("serial")
public class ReportBook implements Serializable{

    /**
     * ワークシート置換情報群
     */
    private List<ReportSheet> reportSheets = new ArrayList<ReportSheet>();

    /**
     * テンプレートファイルURL
     */
    private URL templateFileURL = null;
    
    /**
     * テンプレートファイル名
     */
    private String templateFileName = null;

    /**
     * 出力パス＋ファイル名（拡張子なし）
     */
    private String outputFileName = null;

    /**
     * 出力変換情報
     */
    private ConvertConfiguration[] configurations = null;

    
    /**
     * デフォルトコンストラクタ
     */
    public ReportBook(){
    }
    
    /**
     * @param templateFileName テンプレートファイル名(フルパス)
     * @param outputFileName 出力パス＋ファイル名（拡張子なし）
     * @param configurations 変換情報
     */
    public ReportBook( String templateFileName, String outputFileName, ConvertConfiguration... configurations){
        this.templateFileName = templateFileName;
        this.outputFileName = outputFileName;
        this.configurations = configurations;
    }

    /**
     * @param templateFileName テンプレートファイル名(フルパス)
     * @param outputFileName 出力パス＋ファイル名（拡張子なし）
     * @param formatTypes 変換タイプ
     */
    public ReportBook( String templateFileName, String outputFileName, String... formatTypes){
        this.templateFileName = templateFileName;
        this.outputFileName = outputFileName;

        configurations = new ConvertConfiguration[formatTypes.length];
        for ( int i = 0; i < formatTypes.length; i++) {
            configurations[i] = new ConvertConfiguration( formatTypes[i]);
        }
    }
    
    /**
     * jarに含まれるテンプレートを指定したい場合はこちらを使用する
     * 
     * @param templateFileURL テンプレートファイル
     * @param outputFileName 出力パス＋ファイル名（拡張子なし）
     * @param configurations 変換情報
     */
    public ReportBook( URL templateFileURL, String outputFileName, ConvertConfiguration... configurations){
        this.templateFileURL = templateFileURL;
        this.outputFileName = outputFileName;
        this.configurations = configurations;
    }
    
    /**
     * jarに含まれるテンプレートを指定したい場合はこちらを使用する
     *      
     * @param templateFileURL テンプレートファイル名
     * @param outputFileName 出力パス＋ファイル名（拡張子なし）
     * @param formatTypes 変換タイプ
     */
    public ReportBook( URL templateFileURL, String outputFileName, String... formatTypes){
        this.templateFileURL = templateFileURL;
        this.outputFileName = outputFileName;

        configurations = new ConvertConfiguration[formatTypes.length];
        for ( int i = 0; i < formatTypes.length; i++) {
            configurations[i] = new ConvertConfiguration( formatTypes[i]);
        }
    }

    /**
     * ワークシート置換情報群を設定します。
     * 
     * @param reportSheets ワークシート置換情報群
     */
    public void setReportSheets( List<ReportSheet> reportSheets) {
        this.reportSheets = reportSheets;
    }

    /**
     * ワークシート置換情報群を取得します。
     * 
     * @return ワークシート置換情報群
     */
    public List<ReportSheet> getReportSheets() {
        return reportSheets;
    }

    /**
     * ワークシート置換情報を追加する。
     * 
     * @param reportSheet ワークシート置換情報
     */
    public void addReportSheet( ReportSheet reportSheet) {
        reportSheets.add( reportSheet);
    }

    /**
     * ワークシート置換情報群を追加する。
     * 
     * @param reportSheets ワークシート置換情報群
     */
    public void addReportSheets( List<ReportSheet> reportSheets) {
        this.reportSheets.addAll( reportSheets);
    }

    /**
     * ワークシート置換情報を削除する。
     * 
     * @param reportSheet ワークシート置換情報
     */
    public void removeReportSheet( ReportSheet reportSheet) {
        reportSheets.remove( reportSheet);
    }

    /**
     * ワークシート置換情報群を削除する。
     */
    public void clearReportSheets() {
        reportSheets.clear();
    }

    /**
     * 出力パス＋ファイル名（拡張子なし）を取得します。
     * 
     * @return 出力パス＋ファイル名（拡張子なし）
     */
    public String getOutputFileName() {
        return outputFileName;
    }

    /**
     * 出力パス＋ファイル名（拡張子なし）を設定します。
     * 
     * @param outputFileName 出力パス＋ファイル名（拡張子なし）
     */
    public void setOutputFileName( String outputFileName) {
        this.outputFileName = outputFileName;
    }

    /**
     * 出力変換情報を取得します。
     * 
     * @return 出力変換情報
     */
    public ConvertConfiguration[] getConfigurations() {
        return configurations;
    }

    /**
     * 出力変換情報を設定します。
     * 
     * @param configurations 出力変換情報
     */
    public void setConfigurations( ConvertConfiguration... configurations) {
        this.configurations = configurations;
    }

    /**
     * テンプレートファイル名を取得します。
     * 
     * @return テンプレートファイル名
     */
    public String getTemplateFileName() {
        return templateFileName;
    }

    /**
     * テンプレートファイル名を設定します。
     * 
     * @param templateFileName テンプレートファイル名
     */
    public void setTemplateFileName( String templateFileName) {
        this.templateFileName = templateFileName;
    }
    
    /**
     * テンプレートファイルのURLを取得します。
     * 
     * @return テンプレートファイルのURL
     */
    public URL getTemplateFileURL() {
        return templateFileURL;
    }

    /**
     * テンプレートファイルのURLを設定します。
     * 
     * @param templateFileName テンプレートファイル
     */
    public void setTemplateFileURL( URL templateFileURL) {
        this.templateFileURL = templateFileURL;
    }

}
