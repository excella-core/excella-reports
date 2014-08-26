/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportCreateHelper.java 58 2009-10-19 11:42:04Z a-hoshino $
 * $Revision: 58 $
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
package org.bbreak.excella.reports.processor;

import java.util.HashMap;
import java.util.Map;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.OoPdfExporter;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
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
     * 
     * @return 生成されたデフォルトタグパーサー
     */
    public static Map<String, ReportsTagParser<?>> createDefaultParsers() {

        Map<String, ReportsTagParser<?>> parsers = new HashMap<String, ReportsTagParser<?>>();

        // デフォルト作成
        parsers.put( SingleParamParser.DEFAULT_TAG, new SingleParamParser());
        parsers.put( ImageParamParser.DEFAULT_TAG, new ImageParamParser());
        parsers.put( RowRepeatParamParser.DEFAULT_TAG, new RowRepeatParamParser());
        parsers.put( ColRepeatParamParser.DEFAULT_TAG, new ColRepeatParamParser());
        parsers.put( BlockRowRepeatParamParser.DEFAULT_TAG, new BlockRowRepeatParamParser());
        parsers.put( BlockColRepeatParamParser.DEFAULT_TAG, new BlockColRepeatParamParser());
        parsers.put( SumParamParser.DEFAULT_TAG, new SumParamParser());
        parsers.put( RemoveParamParser.DEFAULT_TAG, new RemoveParamParser());

        return parsers;
    }

    /**
     * デフォルトのエクスポーターを生成する。
     * <li>{@link org.bbreak.excella.reports.exporter.ExcelExporter ExcelExporter}</li>
     * <li>{@link org.bbreak.excella.reports.exporter.OoPdfExporter OoPdfExporter}</li>
     * 
     * @return 生成されたデフォルトエクスポーター
     */
    public static Map<String, ReportBookExporter> createDefaultExporters() {

        Map<String, ReportBookExporter> exporters = new HashMap<String, ReportBookExporter>();

        // デフォルト作成
        exporters.put( ExcelExporter.FORMAT_TYPE, new ExcelExporter());
        exporters.put( OoPdfExporter.FORMAT_TYPE, new OoPdfExporter());

        return exporters;

    }

}
