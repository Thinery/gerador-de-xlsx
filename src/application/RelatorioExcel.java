package application;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.util.HashSet;
import java.util.Set;

public class RelatorioExcel {
    private static final String DIRETORIO_PADRAO = "\\\\montagem\\Produção 2025\\Relatórios"; // Defina o diretório padrão
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(RelatorioExcel::criarInterface);
    }

    private static void criarInterface() {
        JFrame frame = new JFrame("Gerador de Planilha XLSX");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 350);
        
        JPanel panel = new JPanel(new GridLayout(8, 2));
        
        JLabel labelCabecalho = new JLabel("Cabeçalho:");
        String[] opcoesCabecalho = {"Avante", "Face e Fotos", "Face Produções"};
        JComboBox<String> comboCabecalho = new JComboBox<>(opcoesCabecalho);
        
        JLabel labelCidade = new JLabel("Cidade:");
        JTextField campoCidade = new JTextField();
        
        JLabel labelContrato = new JLabel("Contrato:");
        JTextField campoContrato = new JTextField();
        
        JLabel labelProducao = new JLabel("Produção:");
        JTextField campoProducao = new JTextField();
        
        JLabel labelNumeracao = new JLabel("Numeração Inicial-Final:");
        JTextField campoNumeracao = new JTextField();
        
        JLabel labelCancelados = new JLabel("Números Cancelados:");
        JTextField campoCancelados = new JTextField();
        
        JLabel labelTotalFotos = new JLabel("Total de Fotos:");
        JTextField campoTotalFotos = new JTextField();
        
        JButton botaoGerar = new JButton("Gerar Planilha");
        botaoGerar.addActionListener(e -> gerarArquivoXLSX(
                campoCidade.getText(),
                campoContrato.getText(),
                campoProducao.getText(),
                (String) comboCabecalho.getSelectedItem(),
                campoNumeracao.getText(),
                campoCancelados.getText(),
                campoTotalFotos.getText()
        ));
        
        panel.add(labelCabecalho); panel.add(comboCabecalho);
        panel.add(labelCidade); panel.add(campoCidade);
        panel.add(labelContrato); panel.add(campoContrato);
        panel.add(labelProducao); panel.add(campoProducao);
        panel.add(labelNumeracao); panel.add(campoNumeracao);
        panel.add(labelCancelados); panel.add(campoCancelados);
        panel.add(labelTotalFotos); panel.add(campoTotalFotos);
        panel.add(new JLabel()); panel.add(botaoGerar);
        
        frame.add(panel);
        frame.setVisible(true);
    }

    private static void gerarArquivoXLSX(String cidade, String contrato, String producao, String cabecalho, String numeracao, String cancelados, String totalFotos) {
        String caminhoModelo = switch (cabecalho) {
            case "Avante" -> "models\\modeloavante.xlsx";
            case "Face e Fotos" -> "models\\modelofacefotos.xlsx";
            case "Face Produções" -> "models\\modelofaceproducoes.xlsx";
            default -> null;
        };
        
        if (caminhoModelo == null) {
            JOptionPane.showMessageDialog(null, "Erro ao selecionar modelo de planilha.");
            return;
        }
        
        JFileChooser fileChooser = new JFileChooser(DIRETORIO_PADRAO);
        fileChooser.setDialogTitle("Salvar Arquivo");
        fileChooser.setSelectedFile(new File("PlanilhaGerada.xlsx"));
        
        int userSelection = fileChooser.showSaveDialog(null);
        if (userSelection != JFileChooser.APPROVE_OPTION) {
            return;
        }
        
        File arquivoSaida = fileChooser.getSelectedFile();
        
        try (FileInputStream fis = new FileInputStream(caminhoModelo);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            
            sheet.getRow(0).getCell(0).setCellValue("CIDADE: " + cidade);
            sheet.getRow(1).getCell(0).setCellValue("CONTRATO: " + contrato);
            sheet.getRow(2).getCell(0).setCellValue("PRODUÇÃO: " + producao);
            sheet.getRow(0).getCell(7).setCellValue("SEQUÊNCIA: " + numeracao);
            sheet.getRow(1).getCell(7).setCellValue("TOTAL FOTOS: " + totalFotos);
            
            Set<Integer> numerosCancelados = new HashSet<>();
            for (String num : cancelados.split(",")) {
                num = num.trim();
                if (!num.isEmpty() && num.matches("\\d+")) {
                    numerosCancelados.add(Integer.parseInt(num));
                }
            }
            
            String[] partes = numeracao.split("-");
            int inicio = Integer.parseInt(partes[0].trim());
            int fim = Integer.parseInt(partes[1].trim());
            
            int linha = 3;
            int coluna = 0;
            for (int i = inicio; i <= fim; i++) {
                if (linha > 27) {
                    linha = 3;
                    coluna++;
                }
                Row row = sheet.getRow(linha);
                if (row == null) row = sheet.createRow(linha);
                Cell cell = row.getCell(coluna);
                if (cell == null) cell = row.createCell(coluna);
                cell.setCellValue(i);
                if (numerosCancelados.contains(i)) {
                    CellStyle estilo = workbook.createCellStyle();
                    estilo.setFillForegroundColor(IndexedColors.BLACK.getIndex());
                    estilo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(estilo);
                }
                linha++;
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(arquivoSaida)) {
                workbook.write(fileOut);
            }
            
            JOptionPane.showMessageDialog(null, "Arquivo salvo em: " + arquivoSaida.getAbsolutePath());
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Erro ao gerar arquivo: " + e.getMessage());
        }
    }
}
