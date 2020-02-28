/*
 * ------------------------------------------------------------------------
 *
 *  Copyright by KNIME AG, Zurich, Switzerland
 *  Website: http://www.knime.com; Email: contact@knime.com
 *
 *  This program is free software; you can redistribute it and/or modify
 *  it under the terms of the GNU General Public License, Version 3, as
 *  published by the Free Software Foundation.
 *
 *  This program is distributed in the hope that it will be useful, but
 *  WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 *  GNU General Public License for more details.
 *
 *  You should have received a copy of the GNU General Public License
 *  along with this program; if not, see <http://www.gnu.org/licenses>.
 *
 *  Additional permission under GNU GPL version 3 section 7:
 *
 *  KNIME interoperates with ECLIPSE solely via ECLIPSE's plug-in APIs.
 *  Hence, KNIME and ECLIPSE are both independent programs and are not
 *  derived from each other. Should, however, the interpretation of the
 *  GNU GPL Version 3 ("License") under any applicable laws result in
 *  KNIME and ECLIPSE being a combined program, KNIME AG herewith grants
 *  you the additional permission to use and propagate KNIME together with
 *  ECLIPSE with only the license terms in place for ECLIPSE applying to
 *  ECLIPSE and the GNU GPL Version 3 applying for KNIME, provided the
 *  license terms of ECLIPSE themselves allow for the respective use and
 *  propagation of ECLIPSE together with KNIME.
 *
 *  Additional permission relating to nodes for KNIME that extend the Node
 *  Extension (and in particular that are based on subclasses of NodeModel,
 *  NodeDialog, and NodeView) and that only interoperate with KNIME through
 *  standard APIs ("Nodes"):
 *  Nodes are deemed to be separate and independent programs and to not be
 *  covered works.  Notwithstanding anything to the contrary in the
 *  License, the License does not apply to Nodes, you are not required to
 *  license Nodes under the License, and you are granted a license to
 *  prepare and propagate Nodes, in each case even if such Nodes are
 *  propagated with or for interoperation with KNIME.  The owner of a Node
 *  may freely choose the license terms applicable to such Node, including
 *  when such Node is propagated with or for interoperation with KNIME.
 * ---------------------------------------------------------------------
 *
 * History
 *   27.02.2020 (Mareike Hoeger, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi2.node.read4;

import java.io.IOException;
import java.io.InputStream;
import java.lang.ref.WeakReference;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.CancellationException;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.Future;
import java.util.concurrent.atomic.AtomicReference;

import org.apache.commons.lang3.tuple.Triple;
import org.knime.core.data.DataTable;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.util.ViewUtils;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.ext.poi2.node.read4.POIUtils.StopProcessing;
import org.knime.filehandling.core.defaultnodesettings.FileChooserHelper;

/**
 *
 * @author Mareike Hoeger, KNIME GmbH, Konstanz, Germany
 */
public class UpdateTableSwingWorker extends SwingWorkerWithContext<String, Object> {
    private final AtomicReference<DataTable> dt = new AtomicReference<>(null);

    private final XLSReaderNodeDialog m_dialog;

    private String m_fileAccessError;

    private String m_previewMsg;

    private boolean m_clearTableViews;

    private String m_previewTablePanelBorderTitle;

    private boolean m_setReload;

    /**
     * @param dialog
     */
    public UpdateTableSwingWorker(final XLSReaderNodeDialog dialog) {
        m_dialog = dialog;
    }



    @Override
    protected String doInBackgroundWithContext() throws Exception {
        final AtomicReference<DataTable> dt = new AtomicReference<>(null);
        try {

            if (m_dialog.isCurrentlyLoadingNodeSettings()) {
                // do not read from the file while loading settings.
                return null;
            }
            if (!m_dialog.createPreview()) {
                return null;
            }

            List<Path> paths = null;
            try {
                paths = getFileChooserHelper().getPaths();
            } catch (final Exception e1) {
                m_previewMsg = "Could not load file";
            }
            if ((paths == null) || paths.isEmpty()) {
                m_clearTableViews = true;
                return "<no file set>";
            }
            final Path path = paths.get(0);
            final String sheet = m_dialog.getSelectedSheetName();
            if ((sheet == null) || sheet.isEmpty()) {
                return "Could not load file";
            }
            if (sheet == XLSReaderNodeDialog.SCANNING) {
                m_clearTableViews = true;
                return "Loading input file...";
            }

            //To prevent showing invalid content:
            m_clearTableViews = true;
            m_dialog.checkPreviousFuture();
            AtomicReference<CachedExcelTable> currentTableReference = m_dialog.getCurrentTable();
            currentTableReference.set(null);
            m_dialog.setReadRows(-1);

            String localSheet = sheet;
            if (localSheet.equals(XLSReaderNodeDialog.FIRST_SHEET)) {
                localSheet = m_dialog.firstSheetName(path);
                if (localSheet == null) {
                    return "<could not load the file>";
                }
            }
            final XLSUserSettings settings = m_dialog.createSettingsFromComponents();
            try {
                m_previewTablePanelBorderTitle  = "Preview with current settings: [" + localSheet + "]";
                CachedExcelTable table = getSheetTable(path, localSheet, settings.isReevaluateFormulae());
                fileTableSettings(localSheet, settings);
                if (table != null) {
                    currentTableReference.set(table);
                } else {
                    return "<could not load the table>";
                }
                m_dialog.setReadRows(table.lastRow());
                dt.set(table.createDataTable(path, settings, null));
                if (m_dialog.getReadRows() < 0) {
                    m_previewMsg = m_dialog.interruptedMessage();
                }
                return (m_dialog.getReadRows()  >= 0 ? "Content of xls(x) sheet: "
                    : "The first " + -(m_dialog.getReadRows() + 1) + " values of the sheet: ") + localSheet;
            } catch (CancellationException e) {
                m_dialog.checkPreviousFuture();
                m_setReload = true;
                final CachedExcelTable cachedExcelTable = currentTableReference.get();
                if (cachedExcelTable != null) {
                    fileTableSettings(localSheet, settings);
                    dt.set(cachedExcelTable.createDataTable(path, settings, null));
                }
                cancel(false);
                return "<load was interrupted>";
            }
        } catch (final Throwable t) {
            m_dialog.checkPreviousFuture();
            m_setReload = true;
            NodeLogger.getLogger(XLSReaderNodeDialog.class)
                .debug("Unable to create settings for file content view", t);
            m_clearTableViews = true;
            return "<unable to create view>";
        }
    }

    /**
     * @param localSheet
     * @param settings
     */
    private void fileTableSettings(final String localSheet, final XLSUserSettings settings) {
        settings.setHasColHeaders(false);
        settings.setHasRowHeaders(false);
        settings.setKeepXLNames(true);
        settings.setReadAllData(true);
        settings.setFirstRow(0);
        settings.setLastRow(0);
        settings.setFirstColumn(0);
        settings.setLastColumn(0);
        settings.setReevaluateFormulae(false);
        settings.setSkipHiddenColumns(true);
        settings.setSkipEmptyRows(false);
        settings.setSkipHiddenColumns(false);
        settings.setSheetName(localSheet);
    }

    private final FileChooserHelper getFileChooserHelper() throws IOException {
        // timeout is passed from JSpinner, but is only used if Custom or KNIME file system is used
        return new FileChooserHelper(m_dialog.getFSConnection(), m_dialog.getFileChooserSettings().clone(),
            m_dialog.getTimeOut() * 1000);
    }

    /**
     * Should only be called from a background thread as it might load a file.
     *
     * @param file File name.
     * @param sheet Sheet name.
     * @param reevaluateFormulae Reevaluate formulae or not?
     * @return The {@link CachedExcelTable}.
     */
    private CachedExcelTable getSheetTable(final Path path, final String sheet, final boolean reevaluateFormulae) {
        CachedExcelTable sheetTable;
        final Triple<String, String, Boolean> key = Triple.of(path.toString(), sheet, reevaluateFormulae);
        if (!m_sheets.containsKey(key) || ((sheetTable = m_sheets.get(key).get()) == null)) {
            LOGGER.debug("Loading sheet " + sheet + "  of " + path.getFileName().toString());

            try (InputStream stream = Files.newInputStream(path)) {
                checkPreviousFuture();
                final ExecutionMonitor monitor = new ExecutionMonitor();
                monitor.getProgressMonitor().addProgressListener(
                    e -> m_loadingProgress.setValue((int)(100 * e.getNodeProgress().getProgress())));
                final Future<CachedExcelTable> tableFuture = ExcelTableReader.isXlsx(path) && !reevaluateFormulae
                    ? CachedExcelTable.fillCacheFromXlsxStreaming(path, stream, sheet, Locale.ROOT, monitor,
                        m_currentTable)
                    : CachedExcelTable.fillCacheFromDOM(path, stream, sheet, Locale.ROOT, reevaluateFormulae,
                        monitor, m_currentTable);
                checkPreviousFutureAndCancel(m_currentlyRunningFuture.getAndSet(tableFuture));
                ViewUtils.invokeAndWaitInEDT(() -> {
                    m_loadingProgress.setValue(0);
                    m_loadingProgress.setVisible(true);
                    m_cancel.setEnabled(true);
                    m_cancel.setVisible(true);
                });
                sheetTable = tableFuture.get();
                if (!m_currentlyRunningFuture.compareAndSet(tableFuture, null)) {
                    LOGGER.warn("Inconsistency, another thread changed the running future");
                    checkPreviousFuture();
                    m_currentlyRunningFuture.set(null);
                }
                m_sheets.put(key, new WeakReference<>(sheetTable));
            } catch (CancellationException | StopProcessing | InterruptedException | ExecutionException e) {
                sheetTable = m_currentTable.get();
                m_previewUpdateButton.setText(RELOAD);
                fixInterrupt();
                m_currentlyRunningFuture.set(null);
            } catch (final IOException e) {
                throw new RuntimeException(e);
            }
        }
        m_currentTable.set(sheetTable);
        return sheetTable;
    }

    protected void processWithContext(final List<V> chunks) {
    }
    /**
     * {@inheritDoc}
     */
    @Override
    protected void doneWithContext() {
        m_loadingProgress.setVisible(false);
        m_cancel.setVisible(false);

        m_dialog.updatePreviewMessage(m_previewMsg);
        if(m_clearTableViews) {
            m_dialog.clearTableViews();
        }
        m_previewUpdateButton.setText(m_dialog.RELOAD);
        String msg;
        try {
            msg = get();
        } catch (InterruptedException | ExecutionException | CancellationException e) {
            if(e instanceof InterruptedException) {
                //reset interrupt flag
                Thread.currentThread().interrupt();
            }
            msg = "<unable to create view>";
        }

        m_dialog.setFileTablePanelBorderTitle(msg);
        final DataTable newFileDataTable = dt.get();
        m_dialog.setNewFileDataTableInEDT(newFileDataTable);
        m_dialog.getCurrentFileWorker().set(null);
        m_dialog.updatePreviewTable();
    }
}
