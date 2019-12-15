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
