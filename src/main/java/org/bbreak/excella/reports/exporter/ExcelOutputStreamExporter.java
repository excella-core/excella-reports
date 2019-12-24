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
