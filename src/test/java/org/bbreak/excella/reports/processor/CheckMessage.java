/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: CheckMessage.java 5 2009-06-22 07:55:44Z tomo-shibata $
 * $Revision: 5 $
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

public class CheckMessage {
    
    private String message;
    private String expected;
    private String actual;
    
    
    public CheckMessage() {
    }


    public CheckMessage( String message, String expected, String actual) {
        this.message = message;
        this.expected = expected;
        this.actual = actual;
    }


    /**
     * @return the message
     */
    public String getMessage() {
        return message;
    }


    /**
     * @param message the message to set
     */
    public void setMessage( String message) {
        this.message = message;
    }


    /**
     * @return the expected
     */
    public String getExpected() {
        return expected;
    }


    /**
     * @param expected the expected to set
     */
    public void setExpected( String expected) {
        this.expected = expected;
    }


    /**
     * @return the actual
     */
    public String getActual() {
        return actual;
    }


    /**
     * @param actual the actual to set
     */
    public void setActual( String actual) {
        this.actual = actual;
    }

    

}
