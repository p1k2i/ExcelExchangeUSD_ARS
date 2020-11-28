import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;
import java.util.Scanner;

public class ExcelExchangeUSD_ARS_UI extends JFrame {
    private JPanel panelMain;
    private JTextField tfInFile;
    private JButton openButton;
    private JButton loadButton;
    private JComboBox<String> cbSheets;
    private JComboBox<String> cbTables;
    private JComboBox<String> cbColumns;
    private JTextField tfOutFile;
    private JButton pathButton;
    private JComboBox<String> cbConverts;
    private JButton startProcessButton;
    private JRadioButton tableRadioButton;
    private JRadioButton cellsRadioButton;
    private JTextField tfCellBegin;
    private JTextField tfCellEnd;
    private JLabel labelColumn;
    private JTextField httpsWwwDolarsiComTextField;
    private JCheckBox cbDiscardFractional;
    private JCheckBox cbSaveAsText;

    public ExcelExchangeUSD_ARS_UI() {
        openButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
                FileNameExtensionFilter filter = new FileNameExtensionFilter(
                        "Excel file", "xlsx");
                chooser.setFileFilter(filter);
                int returnVal = chooser.showOpenDialog(null);
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    tfInFile.setText(chooser.getSelectedFile().getAbsolutePath());
                    tfOutFile.setText(
                            Utils.addStringToEndFileName(
                                    tfInFile.getText(),
                                    "_"+cbConverts.getSelectedItem().toString()));
                }

                loadButton.getMouseListeners()[1].mouseClicked(e);
            }
        });

        loadButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (tfInFile.getText().isEmpty()) return;
                try {
                    File tempFile = null;
                    try (XSSFWorkbook workbook = Utils.getWorkbookSafe(new File(tfInFile.getText()), tempFile)){
                        cbSheets.removeAllItems();
                        int numberOfSheets = workbook.getNumberOfSheets();
                        for(int sheetIdx = 0; sheetIdx < numberOfSheets; sheetIdx++) {
                            XSSFSheet sheet = workbook.getSheetAt(sheetIdx);
                            cbSheets.addItem(sheet.getSheetName());
                        }
                    } finally {
                        if (tempFile!=null && tempFile.exists()) tempFile.delete();
                    }
                } catch (Throwable exception) {
                    JOptionPane.showMessageDialog(null,
                            exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    exception.printStackTrace();
                }
            }
        });

        pathButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                JFileChooser chooser = new JFileChooser();
                chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int returnVal = chooser.showOpenDialog(null);
                if (returnVal == JFileChooser.APPROVE_OPTION) {
                    tfOutFile.setText(
                            Utils.addStringToEndFileName(
                                    tfInFile.getText(),
                                    "_"+cbConverts.getSelectedItem().toString()));
                }
            }
        });

        startProcessButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                try {
                    ConvertType convertType = ConvertType.valueOf(cbConverts.getSelectedItem().toString());

                    //Instantiating the URL class
                    URL url = new URL("https://www.dolarsi.com/api/api.php?type=valoresprincipales");
                    //Retrieving the contents of the specified page
                    Scanner sc = new Scanner(url.openStream());
                    //Instantiating the StringBuffer class to hold the result
                    StringBuffer sb = new StringBuffer();
                    while(sc.hasNext()) {
                        sb.append(sc.next());
                        //System.out.println(sc.next());
                    }
                    //Retrieving the String from the String Buffer object
                    String responseResult = sb.toString();
                    JSONArray jsonArray = new JSONArray(responseResult);

                    double multiplicator = 1;

                    for (int i=0; i < jsonArray.length(); i++){
                        JSONObject jo = jsonArray.getJSONObject(i);
                        JSONObject joCasa = jo.getJSONObject("casa");

                        String nombre = joCasa.getString("nombre");

                        switch (convertType) {
                            case USD_TO_ARS_SELL:
                                if (nombre.replace(" ","").equals("DolarOficial")){
                                    multiplicator = Double.parseDouble(
                                      joCasa.getString("venta").replace(',','.')
                                    );
                                }
                                break;
                            case USD_TO_ARS_BUY:
                                if (nombre.replace(" ","").equals("DolarOficial")){
                                    multiplicator = Double.parseDouble(
                                            joCasa.getString("compra").replace(',','.')
                                    );
                                }
                                break;
                            case ARS_TO_USD_SELL:
                                if (nombre.replace(" ","").equals("DolarOficial")){
                                    multiplicator = 1d/Double.parseDouble(
                                            joCasa.getString("venta").replace(',','.')
                                    );
                                }
                                break;
                            case ARS_TO_USD_BUY:
                                if (nombre.replace(" ","").equals("DolarOficial")){
                                    multiplicator = 1d/Double.parseDouble(
                                            joCasa.getString("compra").replace(',','.')
                                    );
                                }
                                break;
                            case USD_TO_ARS_SELL_BLUE:
                                if (nombre.replace(" ","").equals("DolarBlue")){
                                    multiplicator = Double.parseDouble(
                                            joCasa.getString("venta").replace(',','.')
                                    );
                                }
                                break;
                            case USD_TO_ARS_BUY_BLUE:
                                if (nombre.replace(" ","").equals("DolarBlue")){
                                    multiplicator = Double.parseDouble(
                                            joCasa.getString("compra").replace(',','.')
                                    );
                                }
                                break;
                            case ARS_TO_USD_SELL_BLUE:
                                if (nombre.replace(" ","").equals("DolarBlue")){
                                    multiplicator = 1d/Double.parseDouble(
                                            joCasa.getString("venta").replace(',','.')
                                    );
                                }
                                break;
                            case ARS_TO_USD_BUY_BLUE:
                                if (nombre.replace(" ","").equals("DolarBlue")){
                                    multiplicator = 1/Double.parseDouble(
                                            joCasa.getString("compra").replace(',','.')
                                    );
                                }
                                break;
                        }
                    }

                    int cellsCountWrote = 0;

                    // create excel
                    File tempFile = null;
                    try (XSSFWorkbook workbook = Utils.getWorkbookSafe(new File(tfInFile.getText()), tempFile)){
                        XSSFSheet sheet = workbook.getSheet(cbSheets.getSelectedItem().toString());

                        final CellReference cellReferenceBegin = new CellReference(tfCellBegin.getText());
                        final CellReference cellReferenceEnd = new CellReference(tfCellEnd.getText());

                        final int rowIndexBegin = cellReferenceBegin.getRow();
                        final int rowIndexEnd = cellReferenceEnd.getRow();
                        final int colIndexBegin = cellReferenceBegin.getCol();
                        final int colIndexEnd = cellReferenceEnd.getCol();

                        final DataFormatter df = new DataFormatter();
                        final boolean discardFrac = cbDiscardFractional.isSelected();
                        final boolean saveAsText = cbSaveAsText.isSelected();

                        for (int rowIndex = rowIndexBegin; rowIndex <= rowIndexEnd; rowIndex++){
                            XSSFRow row = sheet.getRow(rowIndex);
                            try {
                                for (int colIndex = colIndexBegin; colIndex <= colIndexEnd; colIndex++) {
                                    XSSFCell cell = row.getCell(colIndex);
                                    String cellVal = df.formatCellValue(cell);
                                    try {
                                        double cellValD = Double.parseDouble(cellVal) * multiplicator;
                                        if (discardFrac) {
                                            if (saveAsText)
                                                cell.setCellValue(Integer.toString((int) cellValD));
                                            else
                                                cell.setCellValue((int) cellValD);
                                        } else {
                                            if (saveAsText)
                                                cell.setCellValue(Double.toString(cellValD));
                                            else
                                                cell.setCellValue(cellValD);
                                        }
                                        cellsCountWrote++;
                                    } catch (NullPointerException | NumberFormatException ignored) { // no cell exists or is not numeric
                                    }
                                }
                            } catch (NullPointerException ignored){ // no row exists
                            }
                        }

                        File fileOut = new File(tfOutFile.getText());
                        if (fileOut.exists()) fileOut.delete();
                        fileOut.createNewFile();

                        try (FileOutputStream fileOutputStream = new FileOutputStream(fileOut)) {
                            workbook.write(fileOutputStream);
                        }

                    } finally {
                        if (tempFile!=null && tempFile.exists()) tempFile.delete();
                    }

                    JOptionPane.showMessageDialog(null,
                            "Operation complete. "+cellsCountWrote+" cells was changed.",
                            "Done", JOptionPane.INFORMATION_MESSAGE);

                } catch (Throwable exception){
                    JOptionPane.showMessageDialog(null,
                            exception.getClass().getCanonicalName()+
                                    System.lineSeparator()+
                                    exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    exception.printStackTrace();
                }
            }
        });

        cbSheets.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent e) {
                cbTables.removeAllItems();
                if (e.getStateChange()!=ItemEvent.SELECTED || e.getItem()==null || e.getItem().toString().equals("")) return;

                try {
                    File tempFile = null;
                    try (XSSFWorkbook workbook = Utils.getWorkbookSafe(new File(tfInFile.getText()), tempFile)){
                        XSSFSheet sheet = workbook.getSheet(e.getItem().toString());
                        for (XSSFTable table : sheet.getTables()){
                            cbTables.addItem(table.getName());
                        }
                    } finally {
                        if (tempFile!=null && tempFile.exists()) tempFile.delete();
                    }
                } catch (Throwable exception) {
                    JOptionPane.showMessageDialog(null,
                            exception.getClass().getCanonicalName()+
                                    System.lineSeparator()+
                                    exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    exception.printStackTrace();
                }
            }
        });

        cbTables.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent e) {
                cbColumns.removeAllItems();
                if (e.getStateChange()!=ItemEvent.SELECTED || e.getItem()==null || e.getItem().toString().equals("")) return;

                try {
                    File tempFile = null;
                    try (XSSFWorkbook workbook = Utils.getWorkbookSafe(new File(tfInFile.getText()), tempFile)){
                        XSSFSheet sheet = workbook.getSheet(cbSheets.getSelectedItem().toString());
                        for (XSSFTable table : sheet.getTables()){
                            if (!table.getName().equals(e.getItem().toString())) continue;
                            for (int colIndex = table.getStartColIndex(); colIndex<=table.getEndColIndex(); colIndex++){
                                cbColumns.addItem(sheet.getRow(table.getStartRowIndex()).getCell(colIndex).getRichStringCellValue().toString());
                            }
                            break;
                        }
                    } finally {
                        if (tempFile!=null && tempFile.exists()) tempFile.delete();
                    }
                } catch (Throwable exception) {
                    JOptionPane.showMessageDialog(null,
                            exception.getClass().getCanonicalName()+
                                    System.lineSeparator()+
                                    exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    exception.printStackTrace();
                }
            }
        });

        cbColumns.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent e) {
                tfCellBegin.setText("");
                tfCellEnd.setText("");
                if (e.getStateChange()!=ItemEvent.SELECTED || e.getItem()==null || e.getItem().toString().equals("")) return;

                try {
                    File tempFile = null;
                    try (XSSFWorkbook workbook = Utils.getWorkbookSafe(new File(tfInFile.getText()), tempFile)){
                        XSSFSheet sheet = workbook.getSheet(cbSheets.getSelectedItem().toString());
                        for (XSSFTable table : sheet.getTables()){
                            if (!table.getName().equals(cbTables.getSelectedItem().toString())) continue;
                            for (int colIndex = table.getStartColIndex(); colIndex<=table.getEndColIndex(); colIndex++){
                                if (!sheet.getRow(table.getStartRowIndex()).getCell(colIndex).getRichStringCellValue().toString().equals(e.getItem().toString())) continue;

                                try {
                                    tfCellBegin.setText(
                                            sheet.getRow(table.getStartRowIndex()+1).getCell(colIndex).getAddress().toString());
                                    try {
                                        tfCellEnd.setText(
                                                sheet.getRow(Math.min(table.getEndRowIndex(), sheet.getLastRowNum()))
                                                        .getCell(colIndex).getAddress().toString());
                                    } catch (NullPointerException nullPointerException){
                                        tfCellEnd.setText(tfCellBegin.getText());
                                    }
                                } catch (NullPointerException nullPointerException){ // no data in columt
                                    JOptionPane.showMessageDialog(null,
                                            "No data was found in the column", "Warning", JOptionPane.WARNING_MESSAGE);
                                }
                            }
                            break;
                        }
                    } finally {
                        if (tempFile!=null && tempFile.exists()) tempFile.delete();
                    }
                } catch (Throwable exception) {
                    JOptionPane.showMessageDialog(null,
                            exception.getClass().getCanonicalName()+
                                    System.lineSeparator()+
                                    exception.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                    exception.printStackTrace();
                }
            }
        });

        tableRadioButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cbTables.setEnabled(true);
                cbColumns.setEnabled(true);
                tfCellBegin.setEnabled(false);
                tfCellEnd.setEnabled(false);
                labelColumn.setEnabled(true);
            }
        });

        cellsRadioButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                cbTables.setEnabled(false);
                cbColumns.setEnabled(false);
                tfCellBegin.setEnabled(true);
                tfCellEnd.setEnabled(true);
                labelColumn.setEnabled(false);
            }
        });

        cbConverts.addItemListener(new ItemListener() {
            @Override
            public void itemStateChanged(ItemEvent e) {
                if (tfInFile.getText().isEmpty()) return;

                tfOutFile.setText(Utils.addStringToEndFileName(tfInFile.getText(), "_"+e.getItem().toString()));
            }
        });

        for (ConvertType convertType : ConvertType.values()){
            cbConverts.addItem(convertType.name());
        }

        /*
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
         */

        setContentPane(panelMain);
        setTitle("Excel Exchange USD/ARS v1.0.1");
        pack();
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
        this.setLocation(dim.width/2-this.getSize().width/2, dim.height/2-this.getSize().height/2);
        setVisible(true);
    }

    public static void main(String[] args) {
        startUI();
    }

    public static void startUI(){
        ExcelExchangeUSD_ARS_UI ui = new ExcelExchangeUSD_ARS_UI();
    }
}
