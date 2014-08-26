/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ExcelOutputStreamExporter.java 79 2009-10-30 02:28:06Z a-hoshino $
 * $Revision: 79 $
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
 * version 3 along with ExCella Reports.  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.reports.exporter;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * Stream出力エクスポーター
 *
 * @since 1.1
 */
public class ExcelOutputStreamExporter extends ExcelExporter {

    /**
     * 変換タイプ：EXCEL
     */
    public static final String FORMAT_TYPE = "OUTPUT_STREAM_EXCEL";

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ExcelOutputStreamExporter.class);

    /**
     * 出力ストリーム
     */
    private OutputStream outputStream;

    /**
     * コンストラクタ
     *
     * @param outputStream 出力ストリーム
     */
    public ExcelOutputStreamExporter( OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    @Override
    public String getExtention() {
        return "";
    }

    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        try {
            if ( log.isInfoEnabled()) {
                log.info( "処理結果を" + outputStream.getClass().getCanonicalName() + "に出力します");
            }
            book.write( outputStream);
        } catch ( IOException e) {
            throw new ExportException( e);
        }
    }
}
