/*
 * ------------------------------------------------------------------------
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
 * -------------------------------------------------------------------
 *
 * History
 *   Mar 15, 2007 (ohl): created
 */
package org.knime.ext.poi.node.write;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTable;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.StringValue;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.NodeLogger;

/**
 *
 * @author ohl, University of Konstanz
 */
public class XLSWriter {

    private static final NodeLogger LOGGER =
            NodeLogger.getLogger(XLSWriter.class);

    /**
     * Excel (Ver. 2003) can handle datasheets up to 64k x 256 cells!
     */
    private static final int MAX_NUM_OF_ROWS = 65536;

    private static final int MAX_NUM_OF_COLS = 256;

    private final XLSWriterSettings m_settings;

    private final FileOutputStream m_outStream;

    /**
     * Creates a new writer with the specified settings.
     *
     * @param outStream the created workbook will be written to.
     * @param settings the settings.
     */
    public XLSWriter(final FileOutputStream outStream,
            final XLSWriterSettings settings) {
        if (settings == null) {
            throw new NullPointerException("Can't operate with null settings!");
        }
        m_outStream = outStream;
        m_settings = settings;
    }

    /**
     * Writes <code>table</code> with current settings.
     *
     * @param table the table to write to the file
     * @param exec an execution monitor where to check for canceled status and
     *            report progress to. (In case of cancellation, the file will be
     *            deleted.)
     * @throws IOException if any related I/O error occurs
     * @throws CanceledExecutionException if execution in <code>exec</code>
     *             has been canceled
     * @throws NullPointerException if table is <code>null</code>
     */
    public void write(final DataTable table, final ExecutionMonitor exec)
            throws IOException, CanceledExecutionException {

        HSSFWorkbook wb = new HSSFWorkbook();

        int sheetIdx = 0; // in case the table doesn't fit in one sheet
        String sheetName = m_settings.getSheetname();
        if ((sheetName == null) || (sheetName.trim().length() == 0)) {
            sheetName = table.getDataTableSpec().getName();
        }
        // max sheetname length is 32 incl. added running index. We cut it to 25
        if (sheetName.length() > 25) {
            sheetName = sheetName.substring(0, 22) + "...";
        }
        // replace characters like \ / * ? [ ] etc.
        sheetName = replaceInvalidChars(sheetName);

        HSSFSheet sheet = wb.createSheet(sheetName);

        DataTableSpec inSpec = table.getDataTableSpec();
        int numOfCols = inSpec.getNumColumns();
        int rowHdrIncr = m_settings.writeRowID() ? 1 : 0;

        if (numOfCols + rowHdrIncr > MAX_NUM_OF_COLS) {
            LOGGER.warn("The table to write has too many columns! Can't put"
                    + " more than " + MAX_NUM_OF_COLS
                    + " columns in one sheet." + " Truncating columns "
                    + (MAX_NUM_OF_COLS + 1) + " to " + numOfCols);
            numOfCols = MAX_NUM_OF_COLS - rowHdrIncr;
        }
        int numOfRows = -1;
        if (table instanceof BufferedDataTable) {
            numOfRows = ((BufferedDataTable)table).getRowCount();
        }

        int rowIdx = 0; // the index of the row in the XLsheet
        short colIdx = 0; // the index of the cell in the XLsheet

        // write column names
        if (m_settings.writeColHeader()) {

            // Create a new row in the sheet
            HSSFRow hdrRow = sheet.createRow(rowIdx++);

            if (m_settings.writeRowID()) {
                hdrRow.createCell(colIdx++).setCellValue("row ID");
            }
            for (int c = 0; c < numOfCols; c++) {
                String cName = inSpec.getColumnSpec(c).getName();
                hdrRow.createCell(colIdx++).setCellValue(cName);
            }

        } // end of if write column names

        // Guess 80% of the job is generating the sheet, 20% is writing it out
        ExecutionMonitor e = exec.createSubProgress(0.8);

        // write each row of the data
        int rowCnt = 0;
        for (DataRow tableRow : table) {

            colIdx = 0;

            // create a new sheet if the old one is full
            if (rowIdx >= MAX_NUM_OF_ROWS) {
                sheetIdx++;
                sheet = wb.createSheet(sheetName + "(" + sheetIdx + ")");
                rowIdx = 0;
                LOGGER.info("Creating additional sheet to store entire table."
                        + "Additional sheet name: " + sheetName + "("
                        + sheetIdx + ")");
            }

            // set the progress
            String rowID = tableRow.getKey().getString();
            String msg;
            if (numOfRows <= 0) {
                msg = "Writing row " + (rowCnt + 1) + " (\"" + rowID + "\")";
            } else {
                msg =
                        "Writing row " + (rowCnt + 1) + " (\"" + rowID
                                + "\") of " + numOfRows;
                e.setProgress(rowCnt / (double)numOfRows, msg);
            }
            // Check if execution was canceled !
            exec.checkCanceled();

            // Create a new row in the sheet
            HSSFRow sheetRow = sheet.createRow(rowIdx++);

            // add the row id
            if (m_settings.writeRowID()) {
                sheetRow.createCell(colIdx++).setCellValue(rowID);
            }
            // now add all data cells
            for (int c = 0; c < numOfCols; c++) {

                DataCell colValue = tableRow.getCell(c);

                if (colValue.isMissing()) {
                    String miss = m_settings.getMissingPattern();
                    if (miss != null) {
                        sheetRow.createCell(colIdx).setCellValue(miss);
                    }
                } else {
                    HSSFCell sheetCell = sheetRow.createCell(colIdx);

                    if (colValue.getType().isCompatible(DoubleValue.class)) {
                        double val = ((DoubleValue)colValue).getDoubleValue();
                        sheetCell.setCellValue(val);
                    } else if (colValue.getType().isCompatible(
                            StringValue.class)) {
                        String val = ((StringValue)colValue).getStringValue();
                        sheetCell.setCellValue(val);
                    } else {
                        String val = colValue.toString();
                        sheetCell.setCellValue(val);
                    }
                }

                colIdx++;
            }

            rowCnt++;
        } // end of for all rows in table

        // Write the output to a file
        wb.write(m_outStream);

    }

    /**
     * Replaces characters that are illegal in sheet names.
     * These are \/:*?"<>|[].
     *
     * @param name the name to clean
     * @return returns the name with all of the above characters replaced by an
     *         underscore.
     */
    private String replaceInvalidChars(final String name) {
        StringBuilder result = new StringBuilder();
        int l = name.length();
        for (int i = 0; i < l; i++) {
            char c = name.charAt(i);
            if ((c == '\\') || (c == '/') || (c == ':') || (c == '*')
                    || (c == '?') || (c == '"') || (c == '<') || (c == '>')
                    || (c == '|') || (c == '[') || (c == ']')) {
                result.append('_');
            } else {
                result.append(c);
            }
        }
        return result.toString();
    }

}
