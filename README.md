# Document Sender

A web application to upload PDF/JPG/PNG documents and distribute them via Email and SMS to contacts listed in an Excel file.

## Features

- Upload PDF, JPG, or PNG documents
- Add an optional email attachment (documents, spreadsheets, presentations, images, or videos)
- Import contacts from Excel (.xlsx, .xls, .csv) files
- Preview contacts before sending
- Validate email addresses or Indian mobile numbers for the selected channel
- Edit contact details in the Verify Data tab when using the admin role
- Simulated email and SMS sending with progress tracking (email attachments are selected and previewed; actual delivery requires a connected email backend)
- Delivery report after sending

Provider credentials must stay on the backend; do not put provider secrets in `VITE_` variables because Vite exposes those values in the client bundle. Mailtrap is available for a real, allowlisted test email from the Preview & Test tab. Other providers and campaign sends remain simulated. The Mailtrap endpoint is mounted by the Vite development and preview servers for local testing; deploy it behind an authenticated production backend before exposing live sending.

## Excel Format

Your Excel file should have the following columns:

| Name | Email | Mobile |
|------|-------|--------|
| John Doe | john@example.com | +919876543210 |
| Jane Smith | jane@example.com | +919876543211 |

## Setup

```bash
npm install
npm run dev
```

To enable Mailtrap, copy `.env.example` to `.env` in the project root (Vite does not load `.env.example` itself), then replace the placeholder `MAILTRAP_API_TOKEN` and set `MAILTRAP_ALLOWED_FROM_EMAILS` and `MAILTRAP_ALLOWED_RECIPIENTS`. Use an authorized Mailtrap sender address and a recipient present in your uploaded contacts. Both address allowlists accept comma-separated email addresses; sender and recipient values in the app must match them. Restart the Vite server after changing `.env`. The test-send endpoint is limited to five requests per minute per client IP, sends only to the selected preview contact, and supports an optional attachment up to 5 MB. Without a real Mailtrap token, a real email cannot be sent; the app reports the missing configuration instead of simulating Mailtrap success.

## Build

```bash
npm run build
```

## Tech Stack

- React 18 + TypeScript
- Vite
- Tailwind CSS 4
- XLSX (Excel parsing)
- JSZip + FileSaver (ZIP download)
- Lucide React (Icons)

## Backend Integration (Spring Boot + Java 27 + Gradle)

To connect with a Spring Boot backend:

1. Create a Spring Boot application with Java 27
2. Use Gradle as build tool
3. Implement REST API endpoints:
   - POST /api/upload - Upload document
   - POST /api/send-emails - Send emails with attachment
   - POST /api/send-sms - Send SMS with document link
4. Configure JavaMailSender for email
5. Integrate SMS gateway (Twilio, AWS SNS, etc.)

### Sample build.gradle

```groovy
plugins {
    id 'java'
    id 'org.springframework.boot' version '3.4.0'
    id 'io.spring.dependency-management' version '1.1.6'
}

group = 'com.documentsender'
version = '1.0.0'

java {
    toolchain {
        languageVersion = JavaLanguageVersion.of(27)
    }
}

repositories {
    mavenCentral()
}

dependencies {
    implementation 'org.springframework.boot:spring-boot-starter-web'
    implementation 'org.springframework.boot:spring-boot-starter-mail'
    implementation 'org.apache.poi:poi-ooxml:5.2.5'
    implementation 'org.projectlombok:lombok'
    annotationProcessor 'org.projectlombok:lombok'
    testImplementation 'org.springframework.boot:spring-boot-starter-test'
}
```