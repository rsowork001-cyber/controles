package com.nautil;

import com.nautil.admin.config.NautilConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService {

    private static final Logger log = LoggerFactory.getLogger(ExcelService.class);

    @Autowired
    private NautilConfig config;

    private static final int DATA_START_ROW = 3; // Row index 3 = Excel row 4 (0-based)
    private static final int DATA_START_COL_B = 1; // Column B (0-based)
    private static final int DATA_END_COL_G = 6;   // Column G (0-based)
    private static final int DATA_START_COL_A = 0; // Column A (0-based)
    private static final int DATA_END_COL_F = 5;   // Column F (0-based)

    /**
     * Étape 1 : Renommer le fichier de vérification avec la date du jour
     */
    public TaskResult renameVerificationFile() {
        TaskResult result = new TaskResult();
        try {
            String currentPath = config.getVerificationFilePath();
            File currentFile = new File(currentPath);

            if (!currentFile.exists()) {
                result.setSuccess(false);
                result.setMessage("Fichier introuvable : " + currentPath);
                return result;
            }

            String dateStr = LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            String newFileName = currentFile.getName().replaceAll("\\d{4}-\\d{2}-\\d{2}", dateStr);
            File newFile = new File(currentFile.getParent(), newFileName);

            if (currentFile.renameTo(newFile)) {
                config.setVerificationFilePath(newFile.getAbsolutePath());
                result.setSuccess(true);
                result.setMessage("Fichier renommé : " + newFileName);
                result.addLog("✅ Renommage : " + currentFile.getName() + " → " + newFileName);
            } else {
                result.setSuccess(false);
                result.setMessage("Impossible de renommer le fichier");
            }
        } catch (Exception e) {
            log.error("Erreur renommage fichier", e);
            result.setSuccess(false);
            result.setMessage("Erreur : " + e.getMessage());
        }
        return result;
    }

    /**
     * Étape 2 : Ajouter une ligne avec la date du jour dans "Vue Globale - Re7 NAUTIL"
     */
    public TaskResult addDateRowInVueGlobale() {
        TaskResult result = new TaskResult();
        try {
            File file = new File(config.getVerificationFilePath());
            if (!file.exists()) {
                result.setSuccess(false);
                result.setMessage("Fichier introuvable : " + config.getVerificationFilePath());
                return result;
            }

            try (FileInputStream fis = new FileInputStream(file);
                 XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

                String sheetName = config.getSheetGlobal();
                Sheet sheet = workbook.getSheet(sheetName);

                if (sheet == null) {
                    result.setSuccess(false);
                    result.setMessage("Onglet introuvable : " + sheetName);
                    return result;
                }

                // Trouver la première ligne vide après les en-têtes
                int lastRowNum = findLastDataRow(sheet);
                int newRowIndex = lastRowNum + 1;

                Row newRow = sheet.createRow(newRowIndex);
                Cell dateCell = newRow.createCell(0);

                // Créer le style de date
                CellStyle dateStyle = workbook.createCellStyle();
                CreationHelper createHelper = workbook.getCreationHelper();
                dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));

                dateCell.setCellStyle(dateStyle);
                dateCell.setCellValue(java.sql.Date.valueOf(LocalDate.now()));

                result.addLog("✅ Ligne ajoutée dans '" + sheetName + "' à la ligne " + (newRowIndex + 1));
                result.addLog("   Date : " + LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

                saveWorkbook(workbook, file);
                result.setSuccess(true);
                result.setMessage("Ligne de date ajoutée dans Vue Globale");
            }
        } catch (Exception e) {
            log.error("Erreur ajout ligne Vue Globale", e);
            result.setSuccess(false);
            result.setMessage("Erreur : " + e.getMessage());
        }
        return result;
    }

    /**
     * Étape 3 : Supprimer les données de B4:G(fin) dans TX1, TX2, TX3
     */
    public TaskResult clearTxSheets() {
        TaskResult result = new TaskResult();
        List<String> txSheets = new ArrayList<>();
        txSheets.add(config.getSheetTx1());
        txSheets.add(config.getSheetTx2());
        txSheets.add(config.getSheetTx3());

        try {
            File file = new File(config.getVerificationFilePath());
            if (!file.exists()) {
                result.setSuccess(false);
                result.setMessage("Fichier introuvable : " + config.getVerificationFilePath());
                return result;
            }

            try (FileInputStream fis = new FileInputStream(file);
                 XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

                for (String sheetName : txSheets) {
                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet == null) {
                        result.addLog("⚠️ Onglet non trouvé : " + sheetName);
                        continue;
                    }

                    int clearedRows = clearSheetDataRange(sheet, DATA_START_ROW, DATA_START_COL_B, DATA_END_COL_G);
                    result.addLog("✅ Onglet '" + sheetName + "' : " + clearedRows + " ligne(s) vidée(s) (B4:G" + (DATA_START_ROW + clearedRows) + ")");
                }

                saveWorkbook(workbook, file);
                result.setSuccess(true);
                result.setMessage("Données TX1/TX2/TX3 supprimées");
            }
        } catch (Exception e) {
            log.error("Erreur suppression TX sheets", e);
            result.setSuccess(false);
            result.setMessage("Erreur : " + e.getMessage());
        }
        return result;
    }

    /**
     * Étape 4 : Supprimer toutes les données des onglets APRÈS TX3
     */
    public TaskResult clearSheetsAfterTx3() {
        TaskResult result = new TaskResult();

        try {
            File file = new File(config.getVerificationFilePath());
            if (!file.exists()) {
                result.setSuccess(false);
                result.setMessage("Fichier introuvable : " + config.getVerificationFilePath());
                return result;
            }

            try (FileInputStream fis = new FileInputStream(file);
                 XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

                int tx3Index = workbook.getSheetIndex(config.getSheetTx3());
                if (tx3Index < 0) {
                    result.setSuccess(false);
                    result.setMessage("Onglet TX3 non trouvé");
                    return result;
                }

                int totalSheets = workbook.getNumberOfSheets();
                int cleared = 0;

                for (int i = tx3Index + 1; i < totalSheets; i++) {
                    Sheet sheet = workbook.getSheetAt(i);
                    String sheetName = sheet.getSheetName();

                    // Supprimer toutes les lignes de données
                    int lastRow = sheet.getLastRowNum();
                    for (int rowIdx = lastRow; rowIdx >= 0; rowIdx--) {
                        Row row = sheet.getRow(rowIdx);
                        if (row != null) {
                            sheet.removeRow(row);
                        }
                    }
                    cleared++;
                    result.addLog("✅ Onglet '" + sheetName + "' complètement vidé");
                }

                saveWorkbook(workbook, file);
                result.setSuccess(true);
                result.setMessage(cleared + " onglet(s) après TX3 vidé(s)");

                if (cleared == 0) {
                    result.addLog("ℹ️ Aucun onglet après TX3");
                }
            }
        } catch (Exception e) {
            log.error("Erreur suppression onglets après TX3", e);
            result.setSuccess(false);
            result.setMessage("Erreur : " + e.getMessage());
        }
        return result;
    }

    /**
     * Étape 5 : Traitement du fichier template.xlsx - suppression données A4:F(fin)
     */
    public TaskResult clearTemplateFile() {
        TaskResult result = new TaskResult();

        try {
            File file = new File(config.getTemplateFilePath());
            if (!file.exists()) {
                result.setSuccess(false);
                result.setMessage("Template introuvable : " + config.getTemplateFilePath());
                return result;
            }

            try (FileInputStream fis = new FileInputStream(file);
                 XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

                // Traiter le premier onglet (ou tous les onglets selon besoin)
                Sheet sheet = workbook.getSheetAt(0);
                int clearedRows = clearSheetDataRange(sheet, DATA_START_ROW, DATA_START_COL_A, DATA_END_COL_F);

                result.addLog("✅ Template '" + sheet.getSheetName() + "' : " + clearedRows + " ligne(s) vidée(s) (A4:F" + (DATA_START_ROW + clearedRows) + ")");

                saveWorkbook(workbook, file);
                result.setSuccess(true);
                result.setMessage("Template vidé (A4:F" + (DATA_START_ROW + clearedRows) + ")");
            }
        } catch (Exception e) {
            log.error("Erreur suppression template", e);
            result.setSuccess(false);
            result.setMessage("Erreur : " + e.getMessage());
        }
        return result;
    }

    /**
     * Exécuter toutes les étapes Excel en séquence
     */
    public TaskResult runAllExcelTasks() {
        TaskResult finalResult = new TaskResult();
        finalResult.setSuccess(true);

        // Étape 1 : Renommage
        TaskResult step1 = renameVerificationFile();
        finalResult.getLogs().add("=== ÉTAPE 1 : Renommage fichier ===");
        finalResult.getLogs().addAll(step1.getLogs());
        if (!step1.isSuccess()) {
            finalResult.addLog("⚠️ " + step1.getMessage() + " (traitement continué)");
        }

        // Étape 2 : Vue Globale
        TaskResult step2 = addDateRowInVueGlobale();
        finalResult.getLogs().add("=== ÉTAPE 2 : Ajout ligne Vue Globale ===");
        finalResult.getLogs().addAll(step2.getLogs());
        if (!step2.isSuccess()) {
            finalResult.setSuccess(false);
            finalResult.addLog("❌ " + step2.getMessage());
            finalResult.setMessage("Erreur étape 2 : " + step2.getMessage());
            return finalResult;
        }

        // Étape 3 : TX1/TX2/TX3
        TaskResult step3 = clearTxSheets();
        finalResult.getLogs().add("=== ÉTAPE 3 : Nettoyage TX1/TX2/TX3 ===");
        finalResult.getLogs().addAll(step3.getLogs());
        if (!step3.isSuccess()) {
            finalResult.setSuccess(false);
            finalResult.setMessage("Erreur étape 3 : " + step3.getMessage());
            return finalResult;
        }

        // Étape 4 : Onglets après TX3
        TaskResult step4 = clearSheetsAfterTx3();
        finalResult.getLogs().add("=== ÉTAPE 4 : Nettoyage onglets après TX3 ===");
        finalResult.getLogs().addAll(step4.getLogs());
        if (!step4.isSuccess()) {
            finalResult.addLog("⚠️ " + step4.getMessage());
        }

        // Étape 5 : Template
        TaskResult step5 = clearTemplateFile();
        finalResult.getLogs().add("=== ÉTAPE 5 : Nettoyage template.xlsx ===");
        finalResult.getLogs().addAll(step5.getLogs());
        if (!step5.isSuccess()) {
            finalResult.addLog("⚠️ " + step5.getMessage());
        }

        finalResult.setMessage("Traitement Excel terminé avec succès");
        return finalResult;
    }

    // ===================== Méthodes privées utilitaires =====================

    /**
     * Effacer les données dans une plage de colonnes à partir de la ligne startRow
     */
    private int clearSheetDataRange(Sheet sheet, int startRow, int startCol, int endCol) {
        int clearedCount = 0;
        int lastRow = sheet.getLastRowNum();

        for (int rowIdx = startRow; rowIdx <= lastRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null) continue;

            boolean hasData = false;
            for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
                Cell cell = row.getCell(colIdx);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    hasData = true;
                    break;
                }
            }

            if (hasData) {
                for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
                    Cell cell = row.getCell(colIdx);
                    if (cell != null) {
                        cell.setBlank();
                    }
                }
                clearedCount++;
            }
        }

        return clearedCount;
    }

    /**
     * Trouver la dernière ligne avec des données
     */
    private int findLastDataRow(Sheet sheet) {
        int lastRow = sheet.getLastRowNum();
        // Chercher en remontant la dernière ligne non vide
        for (int i = lastRow; i >= 0; i--) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (Cell cell : row) {
                    if (cell != null && cell.getCellType() != CellType.BLANK) {
                        return i;
                    }
                }
            }
        }
        return 0;
    }

    /**
     * Sauvegarder le workbook dans le fichier
     */
    private void saveWorkbook(XSSFWorkbook workbook, File file) throws IOException {
        // Créer le répertoire de sortie si nécessaire
        File outputDir = new File(config.getOutputDir());
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        try (FileOutputStream fos = new FileOutputStream(file)) {
            workbook.write(fos);
        }
    }

    /**
     * Retourner le chemin du fichier de vérification courant
     */
    public String getCurrentVerificationFilePath() {
        return config.getVerificationFilePath();
    }
}
