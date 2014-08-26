/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportProcessAdaptor.java 14 2009-06-24 02:28:00Z tomo-shibata $
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
package org.bbreak.excella.reports.listener;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.SheetData;
import org.bbreak.excella.core.SheetParser;
import org.bbreak.excella.core.exception.ParseException;
import org.bbreak.excella.reports.model.ReportBook;

/**
 * ReportProcessListenerの空実装クラス
 * 
 * @version 1.0
 */
public class ReportProcessAdaptor implements ReportProcessListener {

    public void postParse( Sheet sheet, SheetParser sheetParser, SheetData sheetData) throws ParseException {
    }

    public void preParse( Sheet sheet, SheetParser sheetParser) throws ParseException {
    }

    public void postBookParse( Workbook workbook, ReportBook reportBook) {
    }

    public void preBookParse( Workbook workbook, ReportBook reportBook) {
    }
}
