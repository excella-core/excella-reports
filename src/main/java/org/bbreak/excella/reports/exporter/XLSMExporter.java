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

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.core.util.PoiUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * Excel出力エクスポーター
 * @since 1.0
 */
public class XLSMExporter extends ReportBookExporter {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( XLSMExporter.class);

    /**
     * 変換タイプ：エクセル
     */
    public static final String FORMAT_TYPE = "XLSM";

    /**
     * 拡張子：2007
     */
    public static final String EXTENTION = ".xlsm";

    /*
     * (non-Javadoc)
     * @see org.poireports.exporter.ReportBookExporter#output(org.apache.poi.ss.usermodel.Workbook, org.excelparser.BookData,
     * org.poireports.model.ConvertConfiguration)
     */
    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        if ( !(book instanceof XSSFWorkbook)) {
            throw new IllegalArgumentException( "Workbook is not XSSFWorkbook.");
        }
        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + getFilePath() + "に出力します。");
        }
        try {
            PoiUtil.writeBook( book, getFilePath());
        }
        catch ( Exception e) {
            throw new ExportException( e);
        }
    }

    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    @Override
    public String getExtention() {
        return EXTENTION;
    }

}
