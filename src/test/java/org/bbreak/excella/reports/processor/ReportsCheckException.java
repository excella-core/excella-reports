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
