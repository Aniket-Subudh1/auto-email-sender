const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const fs = require('fs-extra');

class AutoEmailSender {
    constructor(filePath, smtpConfig, emailColumnName, nameColumnName) {
        this.filePath = filePath;
        this.smtpConfig = smtpConfig;
        this.emailColumnName = emailColumnName;  // Name of the column containing the email addresses
        this.nameColumnName = nameColumnName;    // Name of the column containing the names
        this.transporter = nodemailer.createTransport(smtpConfig);
        this.workbook = XLSX.readFile(filePath);
        this.sheetName = this.workbook.SheetNames[0];
        this.worksheet = this.workbook.Sheets[this.sheetName];
        this.data = XLSX.utils.sheet_to_json(this.worksheet);  // Convert sheet to JSON with headers
    }

    getColumnData(columnName) {
        return this.data.map(row => row[columnName]);
    }

    async sendEmailsForColumn(subject, templatePath) {
        const emailColumnData = this.getColumnData(this.emailColumnName);
        const nameColumnData = this.getColumnData(this.nameColumnName);

        for (let i = 0; i < emailColumnData.length; i++) {
            const email = emailColumnData[i];
            const name = nameColumnData[i] || 'User';  // Fallback to 'User' if name is missing
            const replacements = { name };
            const htmlContent = this.createEmailContent(templatePath, replacements);

            try {
                await this.sendEmail(email, subject, htmlContent);
            } catch (error) {
                console.error(`Failed to send email to ${email}. Error: ${error.message}`);
            }
        }
    }

    async sendEmail(to, subject, htmlContent) {
        const mailOptions = {
            from: this.smtpConfig.auth.user,
            to: to,
            subject: subject,
            html: htmlContent,
        };

        try {
            let info = await this.transporter.sendMail(mailOptions);
            console.log(`Email sent successfully to ${to}: ${info.response}`);
        } catch (error) {
            console.error(`Failed to send email to ${to}. Error: ${error.message}`);
        }
    }

    createEmailContent(templatePath, replacements) {
        try {
            let template = fs.readFileSync(templatePath, 'utf-8');
            for (const [key, value] of Object.entries(replacements)) {
                const regex = new RegExp(`{{${key}}}`, 'g');
                template = template.replace(regex, value);
            }
            return template;
        } catch (error) {
            throw new Error(`Error reading template file: ${error.message}`);
        }
    }
}

module.exports = AutoEmailSender;
