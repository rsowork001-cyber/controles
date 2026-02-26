package com.nautil;

import com.nautil.admin.config.NautilConfig;
import com.nautil.admin.service.EmailService;
import com.nautil.admin.service.ExcelService;
import com.nautil.admin.service.TaskResult;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.Map;

@Controller
public class AdminController {

    @Autowired
    private ExcelService excelService;

    @Autowired
    private EmailService emailService;

    @Autowired
    private NautilConfig config;

    /**
     * Page principale d'administration
     */
    @GetMapping("/")
    public String index(Model model) {
        model.addAttribute("currentDate", LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));
        model.addAttribute("verificationFile", config.getVerificationFilePath());
        model.addAttribute("templateFile", config.getTemplateFilePath());
        model.addAttribute("emailTo", config.getDefaultEmailTo());
        return "admin";
    }

    /**
     * API : Exécuter TOUTES les étapes Excel en une fois
     */
    @PostMapping("/api/run-all")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> runAll() {
        Map<String, Object> response = new HashMap<>();
        try {
            TaskResult result = excelService.runAllExcelTasks();
            response.put("success", result.isSuccess());
            response.put("message", result.getMessage());
            response.put("logs", result.getLogs());
        } catch (Exception e) {
            response.put("success", false);
            response.put("message", "Erreur inattendue : " + e.getMessage());
        }
        return ResponseEntity.ok(response);
    }

    /**
     * API : Étape 1 - Renommer le fichier avec la date du jour
     */
    @PostMapping("/api/step/rename")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepRename() {
        return toResponse(excelService.renameVerificationFile());
    }

    /**
     * API : Étape 2 - Ajouter ligne date dans Vue Globale
     */
    @PostMapping("/api/step/vue-globale")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepVueGlobale() {
        return toResponse(excelService.addDateRowInVueGlobale());
    }

    /**
     * API : Étape 3 - Supprimer données TX1/TX2/TX3
     */
    @PostMapping("/api/step/clear-tx")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepClearTx() {
        return toResponse(excelService.clearTxSheets());
    }

    /**
     * API : Étape 4 - Supprimer onglets après TX3
     */
    @PostMapping("/api/step/clear-after-tx3")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepClearAfterTx3() {
        return toResponse(excelService.clearSheetsAfterTx3());
    }

    /**
     * API : Étape 5 - Nettoyer le template
     */
    @PostMapping("/api/step/clear-template")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepClearTemplate() {
        return toResponse(excelService.clearTemplateFile());
    }

    /**
     * API : Étape 6 - Envoyer l'email
     */
    @PostMapping("/api/step/send-email")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> stepSendEmail(
            @RequestParam(required = false) String to,
            @RequestParam(required = false) String subject,
            @RequestParam(required = false) String body) {
        return toResponse(emailService.sendVerificationEmail(to, subject, body));
    }

    /**
     * API : Exécuter Excel + Envoyer email (workflow complet)
     */
    @PostMapping("/api/run-complete")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> runComplete(
            @RequestParam(required = false) String emailTo,
            @RequestParam(required = false) String emailSubject) {

        Map<String, Object> response = new HashMap<>();
        TaskResult excelResult = excelService.runAllExcelTasks();

        if (!excelResult.isSuccess()) {
            response.put("success", false);
            response.put("message", "Erreur traitement Excel : " + excelResult.getMessage());
            response.put("logs", excelResult.getLogs());
            return ResponseEntity.ok(response);
        }

        // Envoi email
        String dateStr = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
        String subject = (emailSubject != null && !emailSubject.isEmpty())
                ? emailSubject
                : "Vérification Quotidienne NAUTIL - " + dateStr;

        TaskResult emailResult = emailService.sendVerificationEmail(emailTo, subject, null);

        // Fusionner les logs
        excelResult.getLogs().add("=== ÉTAPE 6 : Envoi Email ===");
        excelResult.getLogs().addAll(emailResult.getLogs());

        response.put("success", emailResult.isSuccess());
        response.put("message", emailResult.isSuccess()
                ? "Workflow complet terminé avec succès !"
                : "Excel OK, mais erreur email : " + emailResult.getMessage());
        response.put("logs", excelResult.getLogs());

        return ResponseEntity.ok(response);
    }

    /**
     * API : Mettre à jour la configuration
     */
    @PostMapping("/api/config/update")
    @ResponseBody
    public ResponseEntity<Map<String, Object>> updateConfig(
            @RequestParam(required = false) String verificationFile,
            @RequestParam(required = false) String templateFile,
            @RequestParam(required = false) String outputDir) {

        Map<String, Object> response = new HashMap<>();
        if (verificationFile != null && !verificationFile.isEmpty()) {
            config.setVerificationFilePath(verificationFile);
        }
        if (templateFile != null && !templateFile.isEmpty()) {
            config.setTemplateFilePath(templateFile);
        }
        if (outputDir != null && !outputDir.isEmpty()) {
            config.setOutputDir(outputDir);
        }
        response.put("success", true);
        response.put("message", "Configuration mise à jour");
        return ResponseEntity.ok(response);
    }

    // Utilitaire : convertir TaskResult en Map
    private ResponseEntity<Map<String, Object>> toResponse(TaskResult result) {
        Map<String, Object> response = new HashMap<>();
        response.put("success", result.isSuccess());
        response.put("message", result.getMessage());
        response.put("logs", result.getLogs());
        return ResponseEntity.ok(response);
    }
}
