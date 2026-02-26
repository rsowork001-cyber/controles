# Admin — Application de Vérifications Quotidiennes

Application Spring Boot (Java 1.8) avec interface web d'administration pour automatiser
le traitement des fichiers Excel NAUTIL et l'envoi d'emails.

## Prérequis

- Java 8 (JDK 1.8+)
- Maven 3.6+

## Structure du projet

```
nautil-admin/
├── pom.xml
├── README.md
└── src/main/
    ├── java/com/nautil/admin/
    │   ├── NautilAdminApplication.java       ← Point d'entrée Spring Boot
    │   ├── config/
    │   │   └── NautilConfig.java             ← Configuration chemins & email
    │   ├── controller/
    │   │   └── AdminController.java          ← API REST + page web
    │   └── service/
    │       ├── ExcelService.java             ← Manipulation Apache POI
    │       ├── EmailService.java             ← Envoi email JavaMail
    │       └── TaskResult.java               ← Modèle de résultat
    └── resources/
        ├── application.properties            ← ⚠ À CONFIGURER
        └── templates/
            └── admin.html                    ← Interface web admin
```

## Configuration

Éditer `src/main/resources/application.properties` :

```properties
# Chemins des fichiers Excel
nautil.excel.verification-file=C:/NAUTIL/Vérification Quotidiennes Environnement recettes NAUTIL 2026-02-24_V01.xlsx
nautil.excel.template-file=C:/NAUTIL/template.xlsx
nautil.excel.output-dir=C:/NAUTIL/output/

# Email Gmail (utiliser un mot de passe d'application Google)
spring.mail.username=votre_email@gmail.com
spring.mail.password=xxxx_xxxx_xxxx_xxxx   ← Mot de passe d'application Gmail
nautil.email.from=votre_email@gmail.com
nautil.email.default-to=rsowork001@gmail.com
```

### Mot de passe d'application Gmail

Pour utiliser Gmail comme SMTP :
1. Activez l'authentification 2 facteurs sur votre compte Google
2. Allez dans : Mon compte > Sécurité > Mots de passe des applications
3. Générez un mot de passe pour "Mail / Autre"
4. Utilisez ce mot de passe dans `spring.mail.password`

## Compilation & Lancement

```bash
# Compiler
mvn clean package -DskipTests

# Lancer
mvn spring-boot:run

# Ou directement
java -jar target/nautil-admin-1.0.0.jar
```

## Accès

Ouvrir le navigateur : **http://localhost:8080**

## Fonctionnalités

### Workflow Complet (bouton "Lancer le Workflow Complet")
Exécute toutes les étapes d'un coup + envoi email :

1. **Renommage** : Met la date du jour dans le nom du fichier
2. **Vue Globale — Re7 NAUTIL** : Ajoute une ligne avec la date du jour
3. **TX1 / TX2 / TX3** : Supprime les données de B4:G(fin)
4. **Onglets après TX3** : Vide complètement tous les onglets suivants
5. **template.xlsx** : Supprime les données de A4:F(fin)
6. **Email** : Envoie le fichier en pièce jointe à l'adresse configurée

### Étapes Individuelles
Chaque étape peut être exécutée séparément via les boutons correspondants.

### Console de logs
Affiche en temps réel le résultat de chaque opération.

## API REST

| Méthode | URL | Description |
|---------|-----|-------------|
| GET | `/` | Page d'administration |
| POST | `/api/run-all` | Toutes les étapes Excel |
| POST | `/api/run-complete` | Excel + Email |
| POST | `/api/step/rename` | Renommer le fichier |
| POST | `/api/step/vue-globale` | Ajouter ligne Vue Globale |
| POST | `/api/step/clear-tx` | Vider TX1/TX2/TX3 |
| POST | `/api/step/clear-after-tx3` | Vider onglets après TX3 |
| POST | `/api/step/clear-template` | Vider template.xlsx |
| POST | `/api/step/send-email` | Envoyer l'email |

### Paramètres de l'API send-email
```
POST /api/step/send-email?to=email@example.com&subject=Mon objet
```

### Paramètres de run-complete
```
POST /api/run-complete?emailTo=email@example.com&emailSubject=Mon objet
```
