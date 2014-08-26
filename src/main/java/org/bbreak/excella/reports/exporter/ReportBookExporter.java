/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportBookExporter.java 14 2009-06-24 02:28:00Z tomo-shibata $
 * $Revision: 14 $
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
package org.bbreak.excella.reports.exporter;

import java.util.Collections;
import java.util.SortedSet;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.exporter.book.BookExporter;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * 解析結果(ブック)を帳票出力するクラス 指定されたパス(filePath)に出力します。
 * 
 * @since 1.0
 */
public abstract class ReportBookExporter implements BookExporter {

    /**
     * 変換設定情報
     */
    private ConvertConfiguration configuration = null;

    /**
     * 出力先ファイルパス
     */
    private String filePath = null;

    /**
     * 出力先ファイルパスの取得
     * 
     * @return 出力先ファイルパス
     */
    public String getFilePath() {
        return filePath;
    }

    /**
     * 出力先ファイルパスの設定
     * 
     * @param filePath 出力先ファイルパス
     */
    public void setFilePath( String filePath) {
        this.filePath = filePath;
    }

    /*
     * (non-Javadoc)
     * 
     * @see org.excelparser.exporter.book.BookExporter#export(org.apache.poi.ss.usermodel.Workbook, org.excelparser.BookData)
     */
    public void export( Workbook book, BookData bookdata) throws ExportException {
        // テンプレート削除(暫定対応)
        SortedSet<Integer> deleteSheetIndexs = new TreeSet<Integer>( Collections.reverseOrder());
        for ( int i = 0; i < book.getNumberOfSheets(); i++) {
            String sheetName = book.getSheetName( i);
            if ( sheetName.startsWith( PoiUtil.TMP_SHEET_NAME)) {
                deleteSheetIndexs.add( i);
            }
        }
        for ( int index : deleteSheetIndexs) {
            book.removeSheetAt( index);
        }

        if ( configuration != null) {
            // 出力
            output( book, bookdata, configuration);
        }

    }

    /**
     * 出力する。
     * 
     * @param book ワークブック
     * @param bookdata ワークブック解析情報
     * @param configuration 変換設定情報
     * @throws ExportException 出力に失敗した場合
     */
    public abstract void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException;

    /**
     * 変換タイプを取得する。
     * 
     * @return 変換タイプ
     */
    public abstract String getFormatType();
    
    /**
     * 拡張子を取得する。
     * 
     * @return 拡張子
     */
    public abstract String getExtention();
    

    /**
     * 変換設定情報を設定する。
     * 
     * @param configuration 変換設定情報
     */
    public void setConfiguration( ConvertConfiguration configuration) {
        this.configuration = configuration;
    }

    /* (non-Javadoc)
     * @see org.bbreak.excella.core.exporter.book.BookExporter#setup()
     */
    public void setup() {
    }

    /* (non-Javadoc)
     * @see org.bbreak.excella.core.exporter.book.BookExporter#tearDown()
     */
    public void tearDown() {
    }


}
