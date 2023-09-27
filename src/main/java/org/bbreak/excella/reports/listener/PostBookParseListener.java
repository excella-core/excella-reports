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
import org.bbreak.excella.reports.model.ReportBook;

/**
 * ブック処理後に差し込む処理を示すインタフェース
 * 
 * @since 2.2
 */
@FunctionalInterface
public interface PostBookParseListener {

    /**
     * ブック処理後に呼び出されるメソッド
     * 
     * @param workbook 書き込み対象のブック
     * @param reportBook 処理対象の{@link ReportBook}
     */
    void postBookParse( Workbook workbook, ReportBook reportBook);

}
