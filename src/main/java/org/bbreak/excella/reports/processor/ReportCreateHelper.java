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

package org.bbreak.excella.reports.processor;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.BreakParamParser;
import org.bbreak.excella.reports.tag.ColRepeatParamParser;
import org.bbreak.excella.reports.tag.ImageParamParser;
import org.bbreak.excella.reports.tag.RemoveParamParser;
import org.bbreak.excella.reports.tag.ReportsTagParser;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;
import org.bbreak.excella.reports.tag.SumParamParser;

/**
 * 帳票生成を行うヘルパークラス
 * 
 * @since 1.0
 */
public final class ReportCreateHelper {

    /**
     * コンストラクタ
     */
    private ReportCreateHelper() {
    }

    /**
     * デフォルトのタグパーサーを生成する。
     * <li>{@link org.bbreak.excella.reports.tag.SingleParamParser SingleParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.ImageParamParser ImageParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.RowRepeatParamParser RowRepeatParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.ColRepeatParamParser ColRepeatParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.BlockRowRepeatParamParser BlockRowRepeatParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.BlockColRepeatParamParser BlockColRepeatParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.SumParamParser SumParamParser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.RemoveParamParser RemoveParamPaser}</li>
     * <li>{@link org.bbreak.excella.reports.tag.BreakParamParser BreakParamPaser}</li>
     * 
     * @return 生成されたデフォルトタグパーサー
     */
    public static Map<String, ReportsTagParser<?>> createDefaultParsers() {

        Map<String, ReportsTagParser<?>> parsers = new LinkedHashMap<String, ReportsTagParser<?>>();

        // デフォルト作成
        parsers.put( SingleParamParser.DEFAULT_TAG, new SingleParamParser());
        parsers.put( ImageParamParser.DEFAULT_TAG, new ImageParamParser());
        parsers.put( RowRepeatParamParser.DEFAULT_TAG, new RowRepeatParamParser());
        parsers.put( ColRepeatParamParser.DEFAULT_TAG, new ColRepeatParamParser());
        parsers.put( BlockRowRepeatParamParser.DEFAULT_TAG, new BlockRowRepeatParamParser());
        parsers.put( BlockColRepeatParamParser.DEFAULT_TAG, new BlockColRepeatParamParser());
        parsers.put( SumParamParser.DEFAULT_TAG, new SumParamParser());
        parsers.put( RemoveParamParser.DEFAULT_TAG, new RemoveParamParser());
        parsers.put( BreakParamParser.DEFAULT_TAG , new BreakParamParser());

        return parsers;
    }

    /**
     * デフォルトのエクスポーターを生成する。
     * <li>{@link org.bbreak.excella.reports.exporter.ExcelExporter ExcelExporter}</li>
     * 
     * @return 生成されたデフォルトエクスポーター
     */
    public static Map<String, ReportBookExporter> createDefaultExporters() {

        Map<String, ReportBookExporter> exporters = new HashMap<String, ReportBookExporter>();

        // デフォルト作成
        exporters.put( ExcelExporter.FORMAT_TYPE, new ExcelExporter());

        return exporters;

    }

}
