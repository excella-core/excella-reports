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
