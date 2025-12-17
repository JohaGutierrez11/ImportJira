package com.montran;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.File;

public class ExcelReaderApp extends JFrame {

	private static final long serialVersionUID = 1L;
	private JTextField excelFileField, projectNameField, outputCsvField;
    private JButton openFileButton, generateCsvButton;
    private JTextArea outputArea;

    public ExcelReaderApp() {
        setTitle("Generador CSV desde Excel");
        setSize(700, 450);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        // Panel principal con margen y GridBagLayout
        JPanel mainPanel = new JPanel(new GridBagLayout());
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5); // padding

        // Fila 0 - Label y campo de archivo Excel
        gbc.gridx = 0; gbc.gridy = 0;
        gbc.anchor = GridBagConstraints.WEST;
        mainPanel.add(new JLabel("Archivo Excel:"), gbc);

        excelFileField = new JTextField(30);
        gbc.gridx = 1; gbc.gridy = 0;
        gbc.gridwidth = 2;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        mainPanel.add(excelFileField, gbc);

        openFileButton = new JButton("Abrir");
        gbc.gridx = 3; gbc.gridy = 0;
        gbc.gridwidth = 1;
        gbc.fill = GridBagConstraints.NONE;
        mainPanel.add(openFileButton, gbc);

        // Fila 1 - Nombre del Proyecto
        gbc.gridx = 0; gbc.gridy = 1;
        mainPanel.add(new JLabel("Nombre del Proyecto:"), gbc);

        projectNameField = new JTextField(30);
        gbc.gridx = 1; gbc.gridy = 1;
        gbc.gridwidth = 3;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        mainPanel.add(projectNameField, gbc);

        // Fila 2 - Archivo CSV de salida
        gbc.gridx = 0; gbc.gridy = 2;
        mainPanel.add(new JLabel("Archivo CSV de Salida:"), gbc);

        outputCsvField = new JTextField(30);
        gbc.gridx = 1; gbc.gridy = 2;
        gbc.gridwidth = 3;
        mainPanel.add(outputCsvField, gbc);

        // Fila 3 - Botón de generación
        generateCsvButton = new JButton("Generar CSV");
        gbc.gridx = 3; gbc.gridy = 3;
        gbc.gridwidth = 1;
        gbc.anchor = GridBagConstraints.EAST;
        mainPanel.add(generateCsvButton, gbc);

        // Fila 4 - Área de salida
        outputArea = new JTextArea(10, 50);
        outputArea.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(outputArea);
        gbc.gridx = 0; gbc.gridy = 4;
        gbc.gridwidth = 4;
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        mainPanel.add(scrollPane, gbc);

        add(mainPanel);

        // Acción de botón
        openFileButton.addActionListener(e -> openFileDialog());
        generateCsvButton.addActionListener(e -> generateCsv());
    }

    private void openFileDialog() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Archivos Excel", "xlsx"));
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            excelFileField.setText(selectedFile.getAbsolutePath());
        }
    }

    private void generateCsv() {
        String excelFilePath = excelFileField.getText();
        String projectName = projectNameField.getText().trim().replace(" ", "_");
        String outputCsv = outputCsvField.getText().trim();

        if (excelFilePath.isEmpty() || projectName.isEmpty() || outputCsv.isEmpty()) {
            outputArea.setText("Por favor complete todos los campos.");
            return;
        }

        try {
        	
            String advertencia = new Main().processExcelAndGenerateCsv(excelFilePath, projectName, outputCsv);
            String mensaje = "Archivo CSV generado correctamente: " + outputCsv;
            if (!advertencia.isEmpty()) {
                mensaje += "\n\n" + advertencia;
            }

            outputArea.setText(mensaje);
            
        } catch (Exception e) {
            outputArea.setText("Error al generar el CSV: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelReaderApp().setVisible(true));
    }
}
