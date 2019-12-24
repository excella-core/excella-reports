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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * Excel出力エクスポーター
 * 
 * @since 1.0
 */
public class ExcelExporter extends ReportBookExporter {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( ExcelExporter.class);

    /**
     * 変換タイプ：エクセル
     */
    public static final String FORMAT_TYPE = "EXCEL";

    /**
     * 拡張子：〜2003
     */
    public static final String EXTENTION_XLS = ".xls";

    /**
     * 拡張子：2007
     */
    public static final String EXTENTION_XLSX = ".xlsx";
    
    /**
     * 拡張子：2007マクロ有効
     */
    public static final String EXTENTION_XLSM = ".xlsm";

    /**
     * マクロ有効か否か
     */
    private boolean macroAvailable;


    /* (non-Javadoc)
     * @see org.poireports.exporter.ReportBookExporter#output(org.apache.poi.ss.usermodel.Workbook, org.excelparser.BookData, org.poireports.model.ConvertConfiguration)
     */
    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {

        if ( book instanceof HSSFWorkbook) {
            if ( log.isInfoEnabled()) {
                log.info( "XLSExporter呼出");
            }
            XLSExporter exporter = new XLSExporter();
            exporter.setConfiguration( configuration);
            exporter.setFilePath( getFilePath() + XLSExporter.EXTENTION);
            exporter.output( book, bookdata, configuration);
            setFilePath(getFilePath() + XLSExporter.EXTENTION);
        } else if ( book instanceof XSSFWorkbook) {
            if ( log.isInfoEnabled()) {
                log.info( "XLSXExporter呼出");
            }

            if ( macroAvailable) {
                XLSMExporter exporter = new XLSMExporter();
                exporter.setConfiguration( configuration);
                exporter.setFilePath( getFilePath() + XLSMExporter.EXTENTION);
                exporter.output( book, bookdata, configuration);
                setFilePath( getFilePath() + XLSMExporter.EXTENTION);
            }
            else {
                XLSXExporter exporter = new XLSXExporter();
                exporter.setConfiguration( configuration);
                exporter.setFilePath( getFilePath() + XLSXExporter.EXTENTION);
                exporter.output( book, bookdata, configuration);
                setFilePath( getFilePath() + XLSXExporter.EXTENTION);
            }
        }
    }

    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    @Override
    public String getExtention() {
        return "";
    }
    
    /**
     * macroAvailableを取得する。
     * @return macroAvailable
     */
    public boolean isMacroAvailable() {
        return macroAvailable;
    }

    /**
     * macroAvailableを設定する。
     * @param macroAvailable 設定する macroAvailable
     */
    public void setMacroAvailable( boolean macroAvailable) {
        this.macroAvailable = macroAvailable;
    }
}
