package com.montran;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.io.File;

import java.util.List;
import com.montran.model.TestSuite;
import com.montran.service.ExcelSuiteReader;
import com.montran.service.ExcelToCsvService;

import java.util.ArrayList;
import javax.swing.JCheckBox;

public class ExcelReaderApp extends JFrame {

	private static final long serialVersionUID = 1L;
	private JTextField excelFileField, projectNameField, outputCsvField;
	private JButton openFileButton, generateCsvButton;
	private JTextArea outputArea;
	private JPanel suitesPanel;
	private JScrollPane suitesScroll;
	private JButton loadSuitesButton;
	private List<TestSuite> loadedSuites = new ArrayList<>();

	public ExcelReaderApp() {
		setTitle("Generador CSV desde Excel");
		setSize(700, 700);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setLocationRelativeTo(null);

		// Panel principal con margen y GridBagLayout
		JPanel mainPanel = new JPanel(new GridBagLayout());
		mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
		GridBagConstraints gbc = new GridBagConstraints();
		gbc.insets = new Insets(5, 5, 5, 5); // padding

		// Fila 0 - Label y campo de archivo Excel
		gbc.gridx = 0;
		gbc.gridy = 0;
		gbc.anchor = GridBagConstraints.WEST;
		mainPanel.add(new JLabel("Archivo Excel:"), gbc);

		excelFileField = new JTextField(30);
		gbc.gridx = 1;
		gbc.gridy = 0;
		gbc.gridwidth = 2;
		gbc.fill = GridBagConstraints.HORIZONTAL;
		mainPanel.add(excelFileField, gbc);

		openFileButton = new JButton("Abrir");
		gbc.gridx = 3;
		gbc.gridy = 0;
		gbc.gridwidth = 1;
		gbc.fill = GridBagConstraints.NONE;
		mainPanel.add(openFileButton, gbc);

		// Fila 1 - Nombre del Proyecto
		gbc.gridx = 0;
		gbc.gridy = 1;
		mainPanel.add(new JLabel("Nombre del Proyecto:"), gbc);

		projectNameField = new JTextField(30);
		gbc.gridx = 1;
		gbc.gridy = 1;
		gbc.gridwidth = 3;
		gbc.fill = GridBagConstraints.HORIZONTAL;
		mainPanel.add(projectNameField, gbc);

		// Fila 2 - Archivo CSV de salida
		gbc.gridx = 0;
		gbc.gridy = 2;
		mainPanel.add(new JLabel("Archivo CSV de Salida:"), gbc);

		outputCsvField = new JTextField(30);
		gbc.gridx = 1;
		gbc.gridy = 2;
		gbc.gridwidth = 3;
		mainPanel.add(outputCsvField, gbc);

		// Fila 3 - Botones
		
		JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));

		loadSuitesButton = new JButton("Cargar Test Suites");
		generateCsvButton = new JButton("Generar CSV");

		buttonsPanel.add(loadSuitesButton);
		buttonsPanel.add(generateCsvButton);

		gbc.gridx = 0;
		gbc.gridy = 3;
		gbc.gridwidth = 4;
		gbc.fill = GridBagConstraints.NONE;
		gbc.anchor = GridBagConstraints.EAST;

		mainPanel.add(buttonsPanel, gbc);

		// ===== FILA 4 - BLOQUE OUTPUT + SUITES =====
		JPanel bottomBlock = new JPanel(new BorderLayout(0, 10));

		outputArea = new JTextArea(8, 50);
		outputArea.setEditable(false);
		JScrollPane outputScroll = new JScrollPane(outputArea);
		bottomBlock.add(outputScroll, BorderLayout.CENTER);

		suitesPanel = new JPanel();
		suitesPanel.setLayout(new BoxLayout(suitesPanel, BoxLayout.Y_AXIS));
		suitesScroll = new JScrollPane(suitesPanel);
		suitesScroll.setPreferredSize(new Dimension(300, 140));
		bottomBlock.add(suitesScroll, BorderLayout.SOUTH);

		gbc.gridx = 0;
		gbc.gridy = 4;
		gbc.gridwidth = 4;
		gbc.fill = GridBagConstraints.BOTH;
		gbc.weightx = 1;
		gbc.weighty = 1;

		mainPanel.add(bottomBlock, gbc);
		
		add(mainPanel);


		// Acción de botón
		openFileButton.addActionListener(e -> openFileDialog());
		loadSuitesButton.addActionListener(e -> loadSuites());
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

	private void loadSuites() {

        String excelFilePath = excelFileField.getText();

        if (excelFilePath.isEmpty()) {
            outputArea.setText("Seleccione primero un archivo Excel.");
            return;
        }

        try {
            ExcelSuiteReader reader = new ExcelSuiteReader();
            loadedSuites = reader.readSuites(excelFilePath);

            suitesPanel.removeAll();

            for (TestSuite suite : loadedSuites) {
                JCheckBox checkBox = new JCheckBox(suite.getName(), true);
                checkBox.addActionListener(e ->
                        suite.setSelected(checkBox.isSelected()));
                suitesPanel.add(checkBox);
            }

            suitesPanel.revalidate();
            suitesPanel.repaint();

            outputArea.setText(
                    "Test Suites cargadas.\nSeleccione las que desea y luego presione 'Generar CSV'.");

        } catch (Exception e) {
            outputArea.setText("Error al cargar las suites: " + e.getMessage());
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

		if (loadedSuites.isEmpty()) {
			outputArea.setText("Primero debe cargar las Test Suites.");
			return;
		}

		try {
            ExcelToCsvService service = new ExcelToCsvService();

            String advertencia = service.generateCsv(
                    excelFilePath,
                    projectName,
                    outputCsv,
                    loadedSuites
            );

            outputArea.setText("CSV generado correctamente.\n" + advertencia);

        } catch (Exception e) {
            outputArea.setText("Error al generar el CSV: " + e.getMessage());
        }
	}

	public static void main(String[] args) {
		SwingUtilities.invokeLater(() -> new ExcelReaderApp().setVisible(true));
	}
}
