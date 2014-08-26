/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: ReportsCheckException.java 84 2009-10-30 04:07:19Z akira-yokoi $
 * $Revision: 84 $
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
package org.bbreak.excella.reports.processor;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class ReportsCheckException extends Exception {

    private static final long serialVersionUID = -7906436553774376932L;

    private List<CheckMessage> checkMessages = new ArrayList<CheckMessage>();

    public ReportsCheckException( List<CheckMessage> checkMessages) {
        this.checkMessages.addAll( checkMessages);
    }

    /**
     * @param e
     * @return
     * @see java.util.List#add(java.lang.Object)
     */
    public boolean add( CheckMessage e) {
        return checkMessages.add( e);
    }

    /**
     * @param c
     * @return
     * @see java.util.List#addAll(java.util.Collection)
     */
    public boolean addAll( Collection<? extends CheckMessage> c) {
        return checkMessages.addAll( c);
    }

    /**
     * @see java.util.List#clear()
     */
    public void clear() {
        checkMessages.clear();
    }

    /**
     * @return the checkMessages
     */
    public List<CheckMessage> getCheckMessages() {
        return checkMessages;
    }

    public String getCheckMessagesToString() {
        StringBuffer buffer = new StringBuffer();
        if ( !checkMessages.isEmpty()) {
            buffer.append( "相違がありました。\n");
        }
        for ( CheckMessage checkMessage : checkMessages) {
            buffer.append( "[" + checkMessage.getMessage()).append( "]\n");
            buffer.append( "期待値：").append( checkMessage.getExpected()).append( "\n");
            buffer.append( "実測値：").append( checkMessage.getActual()).append( "\n");
        }
        return buffer.toString();
    }
}
