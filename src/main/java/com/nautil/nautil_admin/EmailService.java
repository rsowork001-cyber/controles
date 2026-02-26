package com.nautil;

import com.nautil.admin.config.NautilConfig;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.internet.MimeMessage;
import java.io.File;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

@Service
public class EmailService {

    private static final Logger log = LoggerFactory.getLogger(EmailService.class);

    @Autowired
    private JavaMailSender mailSender;

    @Autowired
    private NautilConfig config;

    /**
     * Envoyer l'email avec le fichier Excel en pièce jointe
     */
    public TaskResult sendVerificationEmail(String toAddress, String customSubject, String customBody) {
        TaskResult result = new TaskResult();

        try {
            String dateStr = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));
            String subject = (customSubject != null && !customSubject.isEmpty())
                    ? customSubject
                    : "Vérification Quotidienne NAUTIL - " + dateStr;

            String body = (customBody != null && !customBody.isEmpty())
                    ? customBody
                    : buildDefaultEmailBody(dateStr);

            String recipient = (toAddress != null && !toAddress.isEmpty())
                    ? toAddress
                    : config.getDefaultEmailTo();

            // Construire le message MIME
            MimeMessage message = mailSender.createMimeMessage();
            MimeMessageHelper helper = new MimeMessageHelper(message, true, "UTF-8");

            helper.setFrom(config.getEmailFrom());
            helper.setTo(recipient);
            helper.setSubject(subject);
            helper.setText(body, true); // true = HTML

            // Pièce jointe : fichier de vérification Excel
            File excelFile = new File(config.getVerificationFilePath());
            if (excelFile.exists()) {
                FileSystemResource fileResource = new FileSystemResource(excelFile);
                helper.addAttachment(excelFile.getName(), fileResource);
                result.addLog("✅ Pièce jointe : " + excelFile.getName());
            } else {
                result.addLog("⚠️ Fichier Excel introuvable : " + config.getVerificationFilePath());
            }

            mailSender.send(message);

            result.setSuccess(true);
            result.setMessage("Email envoyé à " + recipient);
            result.addLog("✅ Email envoyé à : " + recipient);
            result.addLog("   Objet : " + subject);

        } catch (Exception e) {
            log.error("Erreur envoi email", e);
            result.setSuccess(false);
            result.setMessage("Erreur envoi email : " + e.getMessage());
            result.addLog("❌ Échec envoi : " + e.getMessage());
        }

        return result;
    }

    /**
     * Corps de l'email par défaut en HTML
     */
    private String buildDefaultEmailBody(String dateStr) {
        return "<html><body>" +
                "<h2>Vérification Quotidienne NAUTIL</h2>" +
                "<p>Bonjour,</p>" +
                "<p>Veuillez trouver en pièce jointe le fichier de vérification quotidienne " +
                "de l'environnement de recettes NAUTIL pour la date du <strong>" + dateStr + "</strong>.</p>" +
                "<br/>" +
                "<p>Ce rapport a été généré automatiquement par l'application NAUTIL Admin.</p>" +
                "<br/>" +
                "<p>Cordialement,<br/>NAUTIL Admin System</p>" +
                "</body></html>";
    }
}
