/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: OoPdfOutputStreamExporter.java 76 2009-10-29 12:34:20Z a-hoshino $
 * $Revision: 76 $
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

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * PDFをStreamに書き出すExporter
 *
 * @since 1.1
 */
public class OoPdfOutputStreamExporter extends OoPdfExporter {

    /**
     * 一時ファイルのプレフィックス
     */
    private static final String TMP_FILE_PREFIX = "tmp";

    /**
     * 変換タイプ：PDF
     */
    public static final String FORMAT_TYPE = "OUTPUT_STREAM_PDF";

    /**
     * 拡張子：PDF
     */
    public static final String EXTENTION = ".pdf";

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( OoPdfOutputStreamExporter.class);

    /**
     * 出力ストリーム
     */
    private OutputStream outputStream;

    /**
     * コンストラクタ
     *
     * @param outputStream 出力ストリーム
     */
    public OoPdfOutputStreamExporter( OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + outputStream.getClass().getCanonicalName() + "に出力します");
        }

        // 一時的にファイル出力
        int point = getFilePath().indexOf( EXTENTION);
        StringBuffer sb = new StringBuffer( getFilePath());
        sb.insert( point, TMP_FILE_PREFIX);
        String tmpFilePath = sb.toString();
        setFilePath( tmpFilePath);

        // 親クラスでPDFの出力まで実行
        super.output( book, bookdata, configuration);

        File pdfFile = new File( getFilePath());

        // ファイルの内容をStreamに書き出す
        BufferedInputStream in = null;
        BufferedOutputStream out = null;
        try {
            in = new BufferedInputStream( new FileInputStream( pdfFile));
            out = new BufferedOutputStream( outputStream);
            int b;

            while ( (b = in.read()) != -1) {
                out.write( b);
            }
        } catch ( IOException e) {
            throw new ExportException( e);
        } finally {
            try {
                if ( in != null) {
                    in.close();
                }
                if ( out != null) {
                    out.close();
                }
            } catch ( IOException e) {
                throw new ExportException( e);
            } finally {
                // PDFファイルの削除
                pdfFile.delete();
            }
        }
    }
}
