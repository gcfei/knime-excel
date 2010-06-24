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
 *   Apr 8, 2009 (ohl): created
 */
package org.knime.ext.poi.node.read2;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTextField;
import javax.swing.border.Border;
import javax.swing.border.TitledBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.data.DataTable;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.tableview.TableView;
import org.knime.core.node.util.FilesHistoryPanel;
import org.knime.core.node.util.ViewUtils;

/**
 * The dialog to the XLS reader.
 *
 * @author Peter Ohl, KNIME.com, Zurich, Switzerland
 */
public class XLSReaderNodeDialog extends NodeDialogPane {

    private final FilesHistoryPanel m_fileName =
            new FilesHistoryPanel("XLSReader", ".xls");

    private final JComboBox m_sheetName = new JComboBox();

    private final JCheckBox m_hasColHdr = new JCheckBox();

    private final JTextField m_colHdrRow = new JTextField();

    private final JCheckBox m_hasRowIDs = new JCheckBox();

    private final JTextField m_rowIDCol = new JTextField();

    private final JTextField m_firstRow = new JTextField();

    private final JTextField m_lastRow = new JTextField();

    private final JTextField m_firstCol = new JTextField();

    private final JTextField m_lastCol = new JTextField();

    private final JCheckBox m_readAllData = new JCheckBox();

    private final TableView m_fileTable = new TableView();

    private final JPanel m_fileTablePanel = new JPanel();

    private final TableView m_previewTable = new TableView();

    private final JLabel m_errorLabel = new JLabel();

    private final JCheckBox m_skipEmptyCols = new JCheckBox();

    private final JCheckBox m_skipEmptyRows = new JCheckBox();

    private final JCheckBox m_uniquifyRowIDs = new JCheckBox();

    /**
     *
     */
    public XLSReaderNodeDialog() {

        JPanel dlgTab = new JPanel();
        dlgTab.setLayout(new BoxLayout(dlgTab, BoxLayout.Y_AXIS));

        JComponent fileBox = getFileBox();
        fileBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Select file to read:"));
        dlgTab.add(fileBox);

        JPanel settingsBox = new JPanel();
        settingsBox.setLayout(new BoxLayout(settingsBox, BoxLayout.Y_AXIS));
        settingsBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Adjust Settings:"));
        settingsBox.add(getSheetBox());
        settingsBox.add(getColHdrBox());
        settingsBox.add(getRowIDBox());
        settingsBox.add(getAreaBox());
        settingsBox.add(getOptionsBox());
        dlgTab.add(settingsBox);
        dlgTab.add(Box.createVerticalGlue());
        dlgTab.add(Box.createVerticalGlue());
        dlgTab.add(getTablesBox());

        addTab("XLS Reader Settings", dlgTab);

    }

    private JComponent getFileBox() {
        Box fBox = Box.createHorizontalBox();
        fBox.add(Box.createHorizontalGlue());
        m_fileName.addChangeListener(new ChangeListener() {
            public void stateChanged(final ChangeEvent e) {
                fileNameChanged();
            }
        });
        fBox.add(m_fileName);
        fBox.add(Box.createHorizontalGlue());
        return fBox;
    }

    private JComponent getSheetBox() {
        Box sheetBox = Box.createHorizontalBox();
        sheetBox.add(Box.createHorizontalGlue());
        sheetBox.add(new JLabel("Select the sheet to read:"));
        sheetBox.add(Box.createHorizontalStrut(5));
        m_sheetName.setPreferredSize(new Dimension(75, 25));
        m_sheetName.setMaximumSize(new Dimension(75, 25));
        m_sheetName.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                if (e.getStateChange() == ItemEvent.SELECTED) {
                    sheetNameChanged();
                }
            }
        });
        sheetBox.add(m_sheetName);
        sheetBox.add(Box.createHorizontalGlue());
        return sheetBox;
    }

    private void sheetNameChanged() {
        updateFileTable();
        updatePreviewTable();
    }

    private JComponent getColHdrBox() {
        Box colHdrBox = Box.createHorizontalBox();
        colHdrBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Column Names:"));

        m_hasColHdr.setText("Table contains column names in row number:");
        m_hasColHdr.setToolTipText("Enter A, B, C, .... or an index 0 ... 255");
        m_hasColHdr.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_colHdrRow.setPreferredSize(new Dimension(75, 25));
        m_colHdrRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_colHdrRow);

        colHdrBox.add(Box.createHorizontalGlue());
        colHdrBox.add(m_hasColHdr);
        colHdrBox.add(Box.createHorizontalStrut(5));
        colHdrBox.add(m_colHdrRow);
        return colHdrBox;
    }

    private void checkBoxChanged() {
        m_colHdrRow.setEnabled(m_hasColHdr.isSelected());
        m_rowIDCol.setEnabled(m_hasRowIDs.isSelected());
        m_uniquifyRowIDs.setEnabled(m_hasRowIDs.isSelected());
        m_firstCol.setEnabled(!m_readAllData.isSelected());
        m_lastCol.setEnabled(!m_readAllData.isSelected());
        m_firstRow.setEnabled(!m_readAllData.isSelected());
        m_lastRow.setEnabled(!m_readAllData.isSelected());
        updatePreviewTable();
    }

    private JComponent getRowIDBox() {

        Box rowBox = Box.createHorizontalBox();
        m_hasRowIDs.setText("Table contains row IDs in column:");
        m_hasRowIDs.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_rowIDCol.setPreferredSize(new Dimension(75, 25));
        m_rowIDCol.setMaximumSize(new Dimension(75, 25));
        m_rowIDCol.setToolTipText("Enter a number. First row has number 1.");
        addFocusLostListener(m_rowIDCol);
        rowBox.add(Box.createHorizontalGlue());
        rowBox.add(m_hasRowIDs);
        rowBox.add(Box.createHorizontalStrut(5));
        rowBox.add(m_rowIDCol);

        Box uniquifyRowIDBox = Box.createHorizontalBox();
        m_uniquifyRowIDs.setText("Make row IDs unique");
        m_uniquifyRowIDs.setToolTipText("If checked, row IDs are uniquified "
                + "by adding a suffix if necessary (could cause memory "
                + "problems with very large data sets).");
        m_uniquifyRowIDs.setSelected(false);
        m_uniquifyRowIDs.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                updatePreviewTable();
            }
        });
        uniquifyRowIDBox.add(Box.createHorizontalGlue());
        uniquifyRowIDBox.add(m_uniquifyRowIDs);

        Box rowIDBox = Box.createVerticalBox();
        rowIDBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Row IDs:"));
        rowIDBox.add(rowBox);
        rowIDBox.add(uniquifyRowIDBox);
        return rowIDBox;
    }

    private JComponent getAreaBox() {
        Box labelBox = Box.createHorizontalBox();
        labelBox.add(new JLabel("Select the columns and rows to read:"));
        labelBox.add(Box.createHorizontalGlue());

        Box rowsBox = Box.createHorizontalBox();
        m_firstRow.setPreferredSize(new Dimension(75, 25));
        m_firstRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_firstRow);
        m_lastRow.setPreferredSize(new Dimension(75, 25));
        m_lastRow.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_lastRow);
        rowsBox.add(Box.createHorizontalGlue());
        rowsBox.add(new JLabel("read rows from:"));
        rowsBox.add(m_firstRow);
        rowsBox.add(new JLabel("to:"));
        rowsBox.add(m_lastRow);

        Box colsBox = Box.createHorizontalBox();
        m_firstCol.setPreferredSize(new Dimension(75, 25));
        m_firstCol.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_firstCol);
        m_lastCol.setPreferredSize(new Dimension(75, 25));
        m_lastCol.setMaximumSize(new Dimension(75, 25));
        addFocusLostListener(m_lastCol);
        colsBox.add(Box.createHorizontalGlue());
        colsBox.add(new JLabel("read columns from:"));
        colsBox.add(m_firstCol);
        colsBox.add(new JLabel("to:"));
        colsBox.add(m_lastCol);

        Box readAllBox = Box.createHorizontalBox();
        m_readAllData.setText("Read entire data sheet");
        m_readAllData.setToolTipText("If checked, cells that contain "
                + "something (data, format, color, etc.) are read in");
        m_readAllData.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                checkBoxChanged();
            }
        });
        m_readAllData.setSelected(false);
        readAllBox.add(Box.createHorizontalGlue());
        readAllBox.add(m_readAllData);

        Box areaBox = Box.createVerticalBox();
        areaBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Set the Area Of Interest:"));
        areaBox.add(labelBox);
        areaBox.add(colsBox);
        areaBox.add(rowsBox);
        areaBox.add(readAllBox);
        return areaBox;
    }

    private JComponent getOptionsBox() {

        Box skipColsBox = Box.createHorizontalBox();
        m_skipEmptyCols.setText("Skip empty columns");
        m_skipEmptyCols.setToolTipText("If checked, columns that contain "
                + "only missing values are not part of the output table");
        m_skipEmptyCols.setSelected(true);
        m_skipEmptyCols.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                updatePreviewTable();
            }
        });
        skipColsBox.add(Box.createHorizontalGlue());
        skipColsBox.add(m_skipEmptyCols);

        Box skipRowsBox = Box.createHorizontalBox();
        m_skipEmptyRows.setText("Skip empty rows");
        m_skipEmptyRows.setToolTipText("If checked, rows that contain "
                + "only missing values are not part of the output table");
        m_skipEmptyRows.setSelected(true);
        m_skipEmptyRows.addItemListener(new ItemListener() {
            public void itemStateChanged(final ItemEvent e) {
                updatePreviewTable();
            }
        });
        skipRowsBox.add(Box.createHorizontalGlue());
        skipRowsBox.add(m_skipEmptyRows);

        Box optionsBox = Box.createVerticalBox();
        optionsBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "More Options:"));
        optionsBox.add(skipColsBox);
        optionsBox.add(skipRowsBox);
        return optionsBox;

    }

    private void fileNameChanged() {

        String[] names = new String[]{};

        if (m_fileName.getSelectedFile() != null
                && !m_fileName.getSelectedFile().isEmpty()) {
            try {
                names = XLSTable.getSheetNames(m_fileName.getSelectedFile());
            } catch (Exception fnf) {
                // list stays empty
            }
        }
        m_sheetName.setModel(new DefaultComboBoxModel(names));
        if (names.length > 0) {
            m_sheetName.setSelectedIndex(0);
            sheetNameChanged();
        } else {
            m_sheetName.setSelectedIndex(-1);
            clearTableViews();
        }
    }

    private JComponent getTablesBox() {

        JTabbedPane viewTabs = new JTabbedPane();

        m_fileTablePanel.setLayout(new BorderLayout());
        m_fileTablePanel.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "XL Sheet Content:"));
        m_fileTablePanel.add(m_fileTable, BorderLayout.CENTER);

        JPanel previewBox = new JPanel();
        previewBox.setLayout(new BorderLayout());
        previewBox.setBorder(BorderFactory.createTitledBorder(BorderFactory
                .createEtchedBorder(), "Preview with current settings:"));
        previewBox.add(m_previewTable, BorderLayout.CENTER);
        m_errorLabel.setForeground(Color.RED);
        m_errorLabel.setText("");
        Box errBox = Box.createHorizontalBox();
        errBox.add(m_errorLabel);
        errBox.add(Box.createHorizontalGlue());
        errBox.add(Box.createVerticalStrut(30));
        previewBox.add(errBox, BorderLayout.NORTH);
        viewTabs.addTab("Preview", previewBox);
        viewTabs.addTab("File Content", m_fileTablePanel);

        return viewTabs;
    }

    private synchronized void clearTableViews() {
        clearPreview();
        clearFileview();
    }

    private synchronized void clearPreview() {
        ViewUtils.runOrInvokeLaterInEDT(new Runnable() {
            public void run() {
                m_previewTable.setDataTable(null);
            }
        });
    }

    private synchronized void clearFileview() {
        ViewUtils.runOrInvokeLaterInEDT(new Runnable() {
            public void run() {
                m_fileTable.setDataTable(null);
            }
        });
    }

    /**
     * reads the current filename and sheetname and fills the file content view.
     */
    private synchronized void updateFileTable() {

        String file = m_fileName.getSelectedFile();
        if (file == null || file.isEmpty()) {
            setFileTablePanelBorderTitle("<no file set>");
            clearTableViews();
            return;
        }
        String sheet = (String)m_sheetName.getSelectedItem();
        if (sheet == null || sheet.isEmpty()) {
            setFileTablePanelBorderTitle("<no sheet in file>");
            clearTableViews();
            return;
        }

        XLSUserSettings s;
        final DataTable dt;
        try {
            s = createFileviewSettings(file, sheet);
            dt = new XLSTable(s);
            setFileTablePanelBorderTitle("Content of XL sheet: " + sheet);
        } catch (Throwable t) {
            NodeLogger.getLogger(XLSReaderNodeDialog.class).debug(
                    "Unable to create setttings for file content view", t);
            setFileTablePanelBorderTitle("<unable to create view>");
            clearTableViews();
            return;
        }

        ViewUtils.runOrInvokeLaterInEDT(new Runnable() {
            public void run() {
                m_fileTable.setDataTable(dt);
            }
        });

    }

    private void setFileTablePanelBorderTitle(final String title) {
        Border b = m_fileTablePanel.getBorder();
        if (b instanceof TitledBorder) {
            TitledBorder tb = (TitledBorder)b;
            tb.setTitle(title);
        }
    }

    /**
     * reads the current filename and sheetname and fills the preview.
     */
    private synchronized void updatePreviewTable() {

        m_errorLabel.setText("");

        String file = m_fileName.getSelectedFile();
        if (file == null || file.isEmpty()) {
            m_errorLabel.setText("Set a filename.");
            clearTableViews();
            return;
        }
        String sheet = (String)m_sheetName.getSelectedItem();
        if (sheet == null || sheet.isEmpty()) {
            m_errorLabel.setText("No sheets contained in selected file.");
            clearTableViews();
            return;
        }

        XLSUserSettings s;
        final DataTable dt;
        try {
            s = createSettingsFromComponents();
            dt = new XLSTable(s);
        } catch (Throwable t) {
            String msg = t.getMessage();
            if (msg == null || msg.isEmpty()) {
                msg = "no details, sorry.";
            }
            m_errorLabel.setText("Invalid settings: " + msg);
            clearPreview();
            return;
        }

        ViewUtils.runOrInvokeLaterInEDT(new Runnable() {
            public void run() {
                try {
                    m_previewTable.setDataTable(dt);
                } catch ( Throwable t) {
                    m_errorLabel.setText(t.getMessage());
                }
            }
        });

    }

    private XLSUserSettings createSettingsFromComponents()
            throws InvalidSettingsException {
        XLSUserSettings s = new XLSUserSettings();

        s.setFileLocation(m_fileName.getSelectedFile());

        s.setSheetName((String)m_sheetName.getSelectedItem());

        s.setSkipEmptyColumns(m_skipEmptyCols.isSelected());
        s.setSkipEmptyRows(m_skipEmptyRows.isSelected());
        s.setReadAllData(m_readAllData.isSelected());

        s.setHasColHeaders(m_hasColHdr.isSelected());
        try {
            s.setColHdrRow(getNumberFromTextField(m_colHdrRow) - 1);
        } catch (InvalidSettingsException ise) {
            if (m_hasColHdr.isSelected()) {
                throw new InvalidSettingsException("Column Header Row: "
                        + ise.getMessage());
            } // else ignore and don't store an invalid value
        }
        s.setUniquifyRowIDs(m_uniquifyRowIDs.isSelected());
        s.setHasRowHeaders(m_hasRowIDs.isSelected());
        try {
            s.setRowHdrCol(getColumnIndexFromTextField(m_rowIDCol));
        } catch (InvalidSettingsException ise) {
            if (m_hasRowIDs.isSelected()) {
                throw new InvalidSettingsException("Row Header Column Idx: "
                        + ise.getMessage());
            } // else ignore and don't store an invalid value
        }
        try {
            s.setFirstColumn(getColumnIndexFromTextField(m_firstCol));
        } catch (InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("First Column: "
                        + ise.getMessage());
            }
        }
        try {
            s.setLastColumn(getColumnIndexFromTextField(m_lastCol));
        } catch (InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("Last Column: "
                        + ise.getMessage());
            }
        }
        try {
            s.setFirstRow(getNumberFromTextField(m_firstRow) - 1);
        } catch (InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("First Row: "
                        + ise.getMessage());
            }
        }
        try {
            s.setLastRow(getNumberFromTextField(m_lastRow) - 1);
        } catch (InvalidSettingsException ise) {
            if (!m_readAllData.isSelected()) {
                throw new InvalidSettingsException("Last Row: "
                        + ise.getMessage());
            }
        }

        return s;
    }

    /**
     * Creates an int from the specified text field. Throws a ISE if the entered
     * value is empty, is not a number or zero or negative.
     */
    private int getNumberFromTextField(final JTextField t)
            throws InvalidSettingsException {
        String input = t.getText();
        if (input == null || input.isEmpty()) {
            throw new InvalidSettingsException("please enter a number.");
        }
        int i;
        try {
            i = Integer.parseInt(input);
        } catch (NumberFormatException nfe) {
            throw new InvalidSettingsException("not a valid integer number.");
        }
        if (i <= 0) {
            throw new InvalidSettingsException(
                    "number must be larger than zero.");
        }
        return i;
    }

    /**
     * Creates an int from the specified text field. It accepts numbers between
     * 0 and 255 (incl.) or XL column headers (starting at 'A', 'B', ... 'Z',
     * 'AA', etc.) Throws a ISE if the entered value is not valid.
     */
    private int getColumnIndexFromTextField(final JTextField t)
            throws InvalidSettingsException {
        String input = t.getText();
        if (input == null || input.isEmpty()) {
            throw new InvalidSettingsException("please enter a column index.");
        }
        int i;
        try {
            i = Integer.parseInt(input);
        } catch (NumberFormatException nfe) {
            try {
                i = XLSTable.getColumnIndex(input.toUpperCase());
            } catch (IllegalArgumentException iae) {
                throw new InvalidSettingsException(
                        "not a valid column index (enter a number or 'A', 'B', ..., 'Z', 'AA', etc.)");
            }
        }
        if (i < 0) {
            throw new InvalidSettingsException("column index can't negative.");
        }
        if (i > 255) {
            throw new InvalidSettingsException("largest column index is 255");
        }
        return i;
    }

    private void transferSettingsIntoComponents(final XLSUserSettings s) {

        m_fileName.setSelectedFile(s.getFileLocation());

        m_skipEmptyCols.setSelected(s.getSkipEmptyColumns());
        m_skipEmptyRows.setSelected(s.getSkipEmptyRows());
        m_readAllData.setSelected(s.getReadAllData());

        m_hasColHdr.setSelected(s.getHasColHeaders());
        m_colHdrRow.setText("" + (s.getColHdrRow() + 1));
        m_hasRowIDs.setSelected(s.getHasRowHeaders());
        m_uniquifyRowIDs.setSelected(s.getUniquifyRowIDs());

        int val;
        val = s.getRowHdrCol();
        if (val >= 0 && val <= 255) {
            m_rowIDCol.setText(XLSTable.getColLabel(s.getRowHdrCol()));
        } else {
            m_rowIDCol.setText("A");
        }
        val = s.getFirstColumn();
        if (val >= 0 && val <= 255) {
            m_firstCol.setText(XLSTable.getColLabel(s.getFirstColumn()));
        } else {
            m_firstCol.setText("A");
        }
        val = s.getLastColumn();
        if (val >= 0 && val <= 255) {
            m_lastCol.setText(XLSTable.getColLabel(s.getLastColumn()));
        } else {
            m_lastCol.setText("");
        }
        val = s.getFirstRow();
        if (val >= 0) {
            m_firstRow.setText("" + (s.getFirstRow() + 1));
        } else {
            m_firstRow.setText("1");
        }
        val = s.getLastRow();
        if (val >= 0) {
            m_lastRow.setText("" + (s.getLastRow() + 1));
        } else {
            m_lastRow.setText("");
        }

        // clear sheet names
        m_sheetName.setModel(new DefaultComboBoxModel());
        // set new sheet names
        fileNameChanged();
        // select the one
        m_sheetName.setSelectedItem(s.getSheetName());
        sheetNameChanged();

        checkBoxChanged();

    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void saveSettingsTo(final NodeSettingsWO settings)
            throws InvalidSettingsException {
        XLSUserSettings s = createSettingsFromComponents();
        s.save(settings);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    protected void loadSettingsFrom(final NodeSettingsRO settings,
            final DataTableSpec[] specs) throws NotConfigurableException {
        XLSUserSettings s;
        try {
            s = XLSUserSettings.load(settings);
        } catch (InvalidSettingsException e) {
            s = new XLSUserSettings();
        }
        transferSettingsIntoComponents(s);
    }

    /**
     * Sets first and last column/row parameters in the passed settings object.
     * File name and sheet index must be set.
     *
     * @param fileName file to read
     * @param sheetName sheet in that file
     * @return settings for the preview table
     * @throws InvalidSettingsException if file name or sheet index is not
     *             valid.
     * @throws FileNotFoundException if the file is not there
     * @throws IOException if an I/O Error occurred
     */
    private static XLSUserSettings createFileviewSettings(
            final String fileName, final String sheetName)
            throws InvalidSettingsException, FileNotFoundException, IOException {
        if (fileName == null || fileName.isEmpty()) {
            throw new NullPointerException("File location must be set.");
        }
        if (sheetName == null || sheetName.isEmpty()) {
            throw new InvalidSettingsException("Sheet name must be set.");

        }

        XLSUserSettings result = new XLSUserSettings();
        result.setFileLocation(fileName);
        result.setSheetName(sheetName);
        result.setHasColHeaders(false);
        result.setHasRowHeaders(false);
        result.setSkipEmptyColumns(false);
        result.setSkipEmptyRows(false);
        result.setSkipHiddenColumns(false);
        result.setKeepXLColNames(true);
        result.setReadAllData(true);

        XLSTableSettings tableSettings = new XLSTableSettings(result);

        // lets display at least 10 columns
        int colNum =
                tableSettings.getLastColumn() - tableSettings.getFirstColumn()
                        + 1;
        if (colNum < 10) {
            result.setReadAllData(false);
            int incr = 10 - colNum;
            if (tableSettings.getFirstColumn() <= incr) {
                incr -= tableSettings.getFirstColumn();
                result.setFirstColumn(0);
            } else {
                incr -= incr / 2;
                result
                        .setFirstColumn(tableSettings.getFirstColumn() - incr
                                / 2);
            }
            result.setLastColumn(tableSettings.getLastColumn() + incr);
            result.setFirstRow(tableSettings.getFirstRow());
            result.setLastRow(tableSettings.getLastRow());
        }
        return result;
    }

    private void addFocusLostListener(final JTextField field) {
        field.addFocusListener(new FocusListener(){
            public void focusLost(final FocusEvent e) {
                updatePreviewTable();
            }
            public void focusGained(final FocusEvent e) {
                // TODO Auto-generated method stub

            }
        });
    }

}
