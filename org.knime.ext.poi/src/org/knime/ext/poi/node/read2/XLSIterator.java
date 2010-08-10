/*
 * ------------------------------------------------------------------------
 *
 *  Copyright (C) 2003 - 2010
 *  University of Konstanz, Germany and
 *  KNIME GmbH, Konstanz, Germany
 *  Website: http://www.knime.org; Email: contact@knime.org
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
 *  KNIME and ECLIPSE being a combined program, KNIME GMBH herewith grants
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
 *   Apr 27, 2009 (ohl): created
 */
package org.knime.ext.poi.node.read2;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.text.Format;
import java.util.Date;
import java.util.Hashtable;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.concurrent.atomic.AtomicReference;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.IntValue;
import org.knime.core.data.RowKey;
import org.knime.core.data.StringValue;
import org.knime.core.data.container.CloseableRowIterator;
import org.knime.core.data.date.DateAndTimeCell;
import org.knime.core.data.date.DateAndTimeValue;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.util.MutableInteger;

/**
 *
 * @author Peter Ohl, KNIME.com, Zurich, Switzerland
 */
class XLSIterator extends CloseableRowIterator {

    private static final NodeLogger LOGGER =
            NodeLogger.getLogger(XLSIterator.class);

    private final static Integer NOSUFFIX = new Integer(0);

    private final Hashtable<String, Number> m_rowIDhash =
            new Hashtable<String, Number>();

    private final XLSTableSettings m_settings;

    private final DataTableSpec m_spec;

    private BufferedInputStream m_fileStream;

    private final Workbook m_workBook;

    private Sheet m_currentSheet;

    private DataRow m_nextRow;

    private final AtomicReference<RuntimeException> m_exception =
            new AtomicReference<RuntimeException>(null);

    // row count in the xl sheet
    private int m_nextRowIdx;

    // number of rows returned by #next()
    private int m_rowCount;

    /**
     * Iterator for XLS reader table.
     *
     * @param settings with the settings from the analyzer
     * @throws IOException if the file specified in the settings is not
     *             accessible
     * @throws InvalidSettingsException if settings are invalid
     * @throws InvalidFormatException
     */
    XLSIterator(final XLSTableSettings settings) throws IOException,
            InvalidSettingsException, InvalidFormatException {
        m_settings = settings;

        m_spec = m_settings.getDataTableSpec();

        m_fileStream = settings.getBufferedInputStream();
        m_workBook = WorkbookFactory.create(m_fileStream);
        m_currentSheet = m_workBook.getSheet(m_settings.getSheetName());

        m_nextRowIdx = -1;
        setNextRow();
        m_rowCount = 0;

    }

    /**
     * Call this before releasing the last reference to this iterator. It closes
     * the underlying source file. Especially if the iterator didn't run to the
     * end of the table, it is required to call this method. Otherwise the file
     * handle is not released until the garbage collector cleans up. A call to
     * {@link #next()} after disposing of the iterator has undefined behavior.
     */
    @Override
    public void close() {
        closeStream();
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void finalize() throws Throwable {
        closeStream();
        super.finalize();
    }

    private void closeStream() {
        if (m_fileStream != null) {
            try {
                m_fileStream.close();
            } catch (IOException e) {
                // then don't close it
            }
        }
    }

    /**
     * Sets the m_nextRow member (or the m_exception) if there is a next row
     * (otherwise both members are null).
     */
    private void setNextRow() {
        m_nextRow = null;
        m_exception.set(null);
        Row nextXLrow = null;

        if (m_currentSheet == null) {
            closeStream();
            return;
        }

        while (true) {
            m_nextRowIdx++;
            if (m_nextRowIdx > XLSTable.getLastRowIdx(m_currentSheet)) {
                // end of data in the sheet
                closeStream();
                return;
            }
            if (m_nextRowIdx > m_settings.getLastRow()) {
                // beyond range selected by user
                closeStream();
                return;
            }
            if (m_settings.getHasColHeaders()
                    && m_nextRowIdx == m_settings.getColHdrRow()) {
                // skip the column header row
                continue;
            }
            if (m_nextRowIdx < m_settings.getFirstRow()) {
                // skip rows outside the user range
                continue;
            }
            nextXLrow = m_currentSheet.getRow(m_nextRowIdx);
            createDataRow(nextXLrow);
            if (m_nextRow != null || m_exception.get() != null) {
                // createDataRow either sets the nextRow or the exception.
                break;
            }
        }
        return;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public boolean hasNext() {
        // we either have a row for you - or an exception
        return m_nextRow != null || m_exception.get() != null;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public DataRow next() {
        if (!hasNext()) {
            throw new NoSuchElementException(
                    "The XLS iterator proceeded beyond the last row");
        }

        if (m_exception.get() != null) {
            RuntimeException t = m_exception.get();
            // no next row after an exception
            m_nextRow = null;
            m_exception.set(null);
            closeStream();
            throw t;
        }

        DataRow result = m_nextRow;
        // create the next data row for the next iteration
        m_rowCount++;
        setNextRow();
        return result;

    }

    private void createDataRow(final Row xlRow) {

        if (m_settings.getSkipEmptyRows()) {
            if (xlRow == null || xlRow.getFirstCellNum() == -1) {
                return;
            }
        }

        DataCell[] cells = new DataCell[m_spec.getNumColumns()];

        int xlOffset = m_settings.getFirstColumn();
        Set<Integer> skippedCols = m_settings.getSkippedCols();

        for (int colIdx = 0; colIdx < m_spec.getNumColumns(); colIdx++) {

            while (skippedCols.contains(xlOffset)) {
                xlOffset++;
            }

            Cell cell = null;
            if (xlRow != null) {
                // (could be null, if we don't skip empty rows)
                cell = xlRow.getCell(xlOffset);
            }
            DataType t = m_spec.getColumnSpec(colIdx).getType();
            try {
                cells[colIdx] = createDataCellfromXLCell(cell, t, colIdx);
            } catch (RuntimeException e) {
                m_exception.set(e);
                return;
            }
            xlOffset++;
        }

        RowKey key = RowKey.createRowKey(m_rowCount);
        if (m_settings.getKeepXLColNames()) {
            // XL row IDs are just the row numbers (1-based)
            key = new RowKey("" + (m_nextRowIdx + 1));
        } else if (m_settings.getHasRowHeaders()) {
            String xlHdr = getXLRowHdr(xlRow);
            if (xlHdr != null && !xlHdr.isEmpty()) {
                String rowID = xlHdr;
                if (m_settings.getUniquifyRowIDs()) {
                    rowID = uniquifyRowHeader(rowID);
                }
                key = new RowKey(rowID);
            } // if row id it empty keep the default KNIME row id style
        }

        if (m_settings.getSkipEmptyRows()) {
            // skip rows with all missing cells
            boolean isEmpty = true;
            for (DataCell c : cells) {
                if (!c.isMissing()) {
                    isEmpty = false;
                    break;
                }
            }
            if (!isEmpty) {
                m_nextRow = new DefaultRow(key, cells);
            }
        } else {
            m_nextRow = new DefaultRow(key, cells);
        }
    }

    /**
     * Extracts the row ID from the m_nextRow.
     *
     * @return the row ID to the nextRow, of null, if it can't
     */
    private String getXLRowHdr(final Row xlRow) {
        if (!m_settings.getHasRowHeaders() || m_settings.getRowHdrCol() < 0) {
            return null;
        }
        if (xlRow == null) {
            return null;
        }
        if (xlRow.getFirstCellNum() < 0) {
            return null;
        }
        Cell rowID = xlRow.getCell(m_settings.getRowHdrCol());
        if (rowID == null) {
            return null;
        }
        DataCell idCell =
                createDataCellfromXLCell(rowID, StringCell.TYPE, m_settings
                        .getRowHdrCol());
        if (idCell instanceof StringValue) {
            return ((StringValue)idCell).getStringValue();
        } else {
            return idCell.toString();
        }
    }

    private DataCell createDataCellfromXLCell(final Cell cell,
            final DataType expectedType, final int colIdx) {
        if (cell == null) {
            return DataType.getMissingCell();
        }
        // determine the type
        switch (cell.getCellType()) {
        case Cell.CELL_TYPE_BLANK:
            return DataType.getMissingCell();
        case Cell.CELL_TYPE_BOOLEAN:
            boolean b = cell.getBooleanCellValue();
            if (expectedType.isCompatible(StringValue.class)) {
                return new StringCell(Boolean.toString(b));
            } else {
                throw new IllegalStateException(
                        "Invalid cell type in column idx " + colIdx
                                + ", sheet '" + m_settings.getSheetName()
                                + "', row " + cell.getRowIndex());
            }
        case Cell.CELL_TYPE_ERROR:
            LOGGER.warn("Error cell type treated as "
                    + "missing cell in column idx " + colIdx + ", sheet '"
                    + m_settings.getSheetName() + "', row "
                    + cell.getRowIndex());
            return DataType.getMissingCell();
        case Cell.CELL_TYPE_FORMULA:
        case Cell.CELL_TYPE_NUMERIC:
            if (expectedType.isCompatible(DateAndTimeValue.class)) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return new DateAndTimeCell(date.getTime(), true, true,
                            false);
                } else {
                    throw new IllegalStateException(
                            "Invalid cell type in column idx " + colIdx
                                    + " (expected Date), sheet '"
                                    + m_settings.getSheetName() + "', row "
                                    + cell.getRowIndex());
                }
            } else if (expectedType.isCompatible(IntValue.class)) {
                Double num = cell.getNumericCellValue();
                if (new Double(num.intValue()).equals(num)) {
                    return new IntCell(num.intValue());
                } else {
                    throw new IllegalStateException(
                            "Invalid cell type in column idx " + colIdx
                                    + " (is Double, expected Int), sheet '"
                                    + m_settings.getSheetName() + "', row "
                                    + cell.getRowIndex());
                }
            } else if (expectedType.isCompatible(DoubleValue.class)) {
                return new DoubleCell(cell.getNumericCellValue());
            } else if (expectedType.isCompatible(StringValue.class)) {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date d = cell.getDateCellValue();
                    DataFormatter formatter = new DataFormatter();
                    Format cellFormat = formatter.createFormat(cell);
                    return new StringCell(cellFormat.format(d));
                }
                Double num = cell.getNumericCellValue();
                return new StringCell(num.toString());
            } else {
                throw new IllegalStateException(
                        "Invalid cell type in column idx " + colIdx
                                + ", sheet '" + m_settings.getSheetName()
                                + "', row " + cell.getRowIndex());
            }
        case Cell.CELL_TYPE_STRING:
            if (expectedType.isCompatible(StringValue.class)) {
                String s = cell.getRichStringCellValue().getString();
                if (s == null || s.equals(m_settings.getMissValuePattern())) {
                    return DataType.getMissingCell();
                } else {
                    return new StringCell(s);
                }
            } else {
                throw new IllegalStateException(
                        "Invalid cell type in column idx " + colIdx
                                + ", sheet '" + m_settings.getSheetName()
                                + "', row " + cell.getRowIndex());
            }
        default:
            throw new IllegalStateException("Invalid cell type in column idx "
                    + colIdx + ", sheet '" + m_settings.getSheetName()
                    + "', row " + cell.getRowIndex());
        }

    }

    private String uniquifyRowHeader(final String newRowHeader) {

        Number oldSuffix = m_rowIDhash.put(newRowHeader, NOSUFFIX);

        if (oldSuffix == null) {
            // haven't seen the rowID so far.
            return newRowHeader;
        }

        String result = newRowHeader;
        while (oldSuffix != null) {

            // we have seen this rowID before!
            int idx = oldSuffix.intValue();

            assert idx >= NOSUFFIX.intValue();

            idx++;

            if (oldSuffix == NOSUFFIX) {
                // until now the NOSUFFIX placeholder was in the hash
                assert idx - 1 == NOSUFFIX.intValue();
                m_rowIDhash.put(result, new MutableInteger(idx));
            } else {
                assert oldSuffix instanceof MutableInteger;
                ((MutableInteger)oldSuffix).inc();
                assert idx == oldSuffix.intValue();
                // put back the old (incr.) suffix (overridden with NOSUFFIX).
                m_rowIDhash.put(result, oldSuffix);
            }

            result = result + "_" + idx;
            oldSuffix = m_rowIDhash.put(result, NOSUFFIX);

        }

        return result;
    }

}
