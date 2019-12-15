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
