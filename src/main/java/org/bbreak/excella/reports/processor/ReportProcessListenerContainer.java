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
import java.util.List;
import java.util.function.Consumer;

import org.bbreak.excella.core.listener.PostSheetParseListener;
import org.bbreak.excella.core.listener.PreSheetParseListener;
import org.bbreak.excella.reports.listener.PostBookParseListener;
import org.bbreak.excella.reports.listener.PreBookParseListener;
import org.bbreak.excella.reports.listener.ReportProcessListener;

class ReportProcessListenerContainer {

    private final List<PreBookParseListener> preBookParseListeners = new ArrayList<>();

    private final List<PreSheetParseListener> preParseListeners = new ArrayList<>();

    private final List<PostSheetParseListener> postParseListeners = new ArrayList<>();

    private final List<PostBookParseListener> postBookParseListeners = new ArrayList<>();

    void addListener( ReportProcessListener listener) {
        addPreBookParseListener( listener);
        addPreSheetParseListener( listener);
        addPostSheetParseListener( listener);
        addPostBookParseListener( listener);
    }

    void removeListener( ReportProcessListener listener) {
        removePreBookParseListener( listener);
        removePreSheetParseListener( listener);
        removePostSheetParseListener( listener);
        removePostBookParseListener( listener);
    }
    void clearListeners() {
        clearPreBookParseListeners();
        clearPreSheetParseListeners();
        clearPostSheetParseListeners();
        clearPostBookParseListeners();
    }

    void addPreBookParseListener( PreBookParseListener listener) {
        preBookParseListeners.add( listener);
    }

    void addPreSheetParseListener( PreSheetParseListener listener) {
        preParseListeners.add( listener);
    }

    void addPostSheetParseListener( PostSheetParseListener listener) {
        postParseListeners.add( listener);
    }

    void addPostBookParseListener( PostBookParseListener listener) {
        postBookParseListeners.add( listener);
    }

    void removePreBookParseListener( PreBookParseListener listener) {
        preBookParseListeners.remove( listener);
    }

    void removePreSheetParseListener( PreSheetParseListener listener) {
        preParseListeners.remove( listener);
    }

    void removePostSheetParseListener( PostSheetParseListener listener) {
        postParseListeners.remove( listener);
    }

    void removePostBookParseListener( PostBookParseListener listener) {
        postBookParseListeners.remove( listener);
    }
    void forEachPreBookParseListeners( Consumer<PreBookParseListener> consumer) {
        preBookParseListeners.forEach( consumer);
    }

    void forEachPreSheetParseListeners( Consumer<PreSheetParseListener> consumer) {
        preParseListeners.forEach( consumer);
    }

    void forEachPostSheetParseListeners( Consumer<PostSheetParseListener> consumer) {
        postParseListeners.forEach( consumer);
    }

    void forEachPostBookParseListeners( Consumer<PostBookParseListener> consumer) {
        postBookParseListeners.forEach( consumer);
    }

    void clearPreBookParseListeners() {
        preBookParseListeners.clear();
    }

    void clearPreSheetParseListeners() {
        preParseListeners.clear();
    }

    void clearPostSheetParseListeners() {
        postParseListeners.clear();
    }

    void clearPostBookParseListeners() {
        postBookParseListeners.clear();
    }

}
