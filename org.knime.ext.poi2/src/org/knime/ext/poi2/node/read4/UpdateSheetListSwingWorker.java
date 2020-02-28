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
 *   25.02.2020 (Mareike Hoeger, KNIME GmbH, Konstanz, Germany): created
 */
package org.knime.ext.poi2.node.read4;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutionException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.knime.core.node.NodeLogger;
import org.knime.core.util.SwingWorkerWithContext;
import org.knime.filehandling.core.defaultnodesettings.FileChooserHelper;

/**
 *
 * @author Mareike Hoeger, KNIME GmbH, Konstanz, Germany
 */
public class UpdateSheetListSwingWorker extends SwingWorkerWithContext<String[], Object> {
    private final XLSReaderNodeDialog m_dialog;

    long m_currentId;

    private String m_fileAccessError;

    private String m_previewMsg;

    private final String m_sheetName;
    /**
     * @param currentId
     * @param xlsReaderNodeDialog
     */
    public UpdateSheetListSwingWorker(final long currentId, final XLSReaderNodeDialog xlsReaderNodeDialog, final String sheetName) {
        m_dialog = xlsReaderNodeDialog;
        m_currentId = currentId;
        m_sheetName = sheetName;
    }

    @Override
    protected String[] doInBackgroundWithContext() throws Exception {
        final List<Path> paths = getFileChooserHelper().getPaths();
        if (paths != null && !paths.isEmpty()) {
            final Path path = paths.get(0);
            final String file = path.toString();
            Workbook workbook = m_dialog.getWorkbook();
            if ((workbook == null) && !ExcelTableReader.isXlsx(file)) {
                workbook = m_dialog.createAndGetWorkbook(path);
            }
            if (workbook != null) {
                try {
                    m_fileAccessError = null;
                    final ArrayList<String> sheetNames = POIUtils.getSheetNames(workbook);
                    sheetNames.add(0, XLSReaderNodeDialog.FIRST_SHEET);
                    return sheetNames.toArray(new String[sheetNames.size()]);
                } catch (final Exception fnf) {
                    NodeLogger.getLogger(XLSReaderNodeDialog.class).error(fnf.getMessage(), fnf);
                    m_fileAccessError = fnf.getMessage();
                    // return empty list then
                }
            } else {//xlsx without reevaluation
                final List<String> sheetNames =
                    POIUtils.getSheetNames(new XSSFReader(OPCPackage.open(Files.newInputStream(path))));
                sheetNames.add(0, XLSReaderNodeDialog.FIRST_SHEET);
                return sheetNames.stream().toArray(n -> new String[n]);
            }
        } else {
            m_previewMsg = "No input file available";
        }
        return new String[]{};
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void doneWithContext() {
        if (m_currentId != m_dialog.getUpdateSheetListId().get()) {
            // Another update of the sheet list has started
            // Do not update the sheet list
            return;
        }
        m_dialog.updatePreviewMessage(m_previewMsg);
        m_dialog.setFileAccessError(m_fileAccessError);
        String[] names = new String[]{};
        try {
            names = get();
        } catch (InterruptedException ie) {
            //reset interrupt flag
            Thread.currentThread().interrupt();
        } catch (ExecutionException e) {

            // ignore
        }
        m_dialog.updateSheetNameCombo(names, m_sheetName);
    }

    private final FileChooserHelper getFileChooserHelper() throws IOException {
        // timeout is passed from JSpinner, but is only used if Custom or KNIME file system is used
        return new FileChooserHelper(m_dialog.getFSConnection(), m_dialog.getFileChooserSettings().clone(),
            m_dialog.getTimeOut() * 1000);
    }

    /**
     * Loads a workbook from the file system.
     *
     * @param path Path to the workbook
     * @return The workbook or null if it could not be loaded
     * @throws IOException
     * @throws InvalidFormatException
     * @throws RuntimeException the underlying POI library also throws other kind of exceptions
     */
    public Workbook getWorkbook(final Path path) throws IOException, InvalidFormatException {
        try (InputStream in = Files.newInputStream(path)) {
            // This should be the only place in the code where a workbook gets loaded
            return WorkbookFactory.create(in);
        }
    }
}
