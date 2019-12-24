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

public class CellObject {
    
    private int rowIndex = -1;

    private int colIndex = -1;

    public CellObject( int rowIndex, int colIndex) {
        this.rowIndex = rowIndex;
        this.colIndex = colIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex( int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex( int colIndex) {
        this.colIndex = colIndex;
    }

    @Override
    public boolean equals( Object obj) {
        CellObject cell = ( CellObject) obj;
        return rowIndex == cell.rowIndex && colIndex == cell.colIndex;
    }

    @Override
    public String toString() {
        
        return "(" + rowIndex + ","  + colIndex + ")";
    }

    
}
