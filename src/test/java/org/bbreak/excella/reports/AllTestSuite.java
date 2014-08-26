/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: AllTestSuite.java 39 2009-09-29 03:38:36Z tsuchida $
 * $Revision: 39 $
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
package org.bbreak.excella.reports;

import org.bbreak.excella.reports.exporter.ExporterTestSuite;
import org.bbreak.excella.reports.listener.ListenerTestSuite;
import org.bbreak.excella.reports.model.ModelTestSuite;
import org.bbreak.excella.reports.others.OthersTestSuite;
import org.bbreak.excella.reports.processor.ProcessorTestSuite;
import org.bbreak.excella.reports.tag.TagTestSuite;
import org.bbreak.excella.reports.util.UtilTestSuite;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

/**
 * 全テスト実行用テストスイート
 * 
 * @since 1.0
 */
@RunWith(Suite.class)
@SuiteClasses({
    TagTestSuite.class,
    UtilTestSuite.class,
    ExporterTestSuite.class,
    ProcessorTestSuite.class,
    ModelTestSuite.class,
    ListenerTestSuite.class,
    OthersTestSuite.class
})
public class AllTestSuite {
    
}
