/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportProcessListener.java 5 2009-06-22 07:55:44Z tomo-shibata $
 * $Revision: 5 $
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
package org.bbreak.excella.reports.listener;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.listener.SheetParseListener;
import org.bbreak.excella.reports.model.ReportBook;

/**
 * レポート作成処理の下記のタイミングで任意の処理を差し込むためのインターフェイス
 * <ul>
 * <li>ブック処理前</li>
 * <li>シート処理前(パラメータ置換前)</li>
 * <li>シート処理後(パラメータ置換後)</li>
 * <li>ブック処理後</li>
 * </ul>
 * 
 * @version 1.0
 */
public interface ReportProcessListener extends SheetParseListener {
    /**
     * ブック処理後に呼び出されるメソッド
     * 
     * @param workbook 書き込み対象のブック
     */
    void preBookParse( Workbook workbook, ReportBook reportBook);

    /**
     * ブック処理後に呼び出されるメソッド
     * 
     * @param workbook 書き込み対象のブック
     */
    void postBookParse( Workbook workbook, ReportBook reportBook);
}
