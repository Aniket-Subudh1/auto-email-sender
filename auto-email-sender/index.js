const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const fs = require('fs-extra');

class AutoEmailSender {
    constructor(filePath, smtpConfig, emailColumnIndex) {
        this.filePath = filePath;
        this.smtpConfig = smtpConfig;
        this.emailColumnIndex = emailColumnIndex;  // New parameter for email column index
        this.transporter = nodemailer.createTransport(smtpConfig);
        this.workbook = XLSX.readFile(filePath);
        this.sheetName = this.workbook.SheetNames[0];
        this.worksheet = this.workbook.Sheets[this.sheetName];
        this.data = XLSX.utils.sheet_to_json(this.worksheet, { header: 1 });
    }

    getRow(rowIndex) {
        if (rowIndex < 0 || rowIndex >= this.data.length) {
            throw new Error('Row index out of range.');
        }
        const row = this.data[rowIndex];

        // Use the specified column index to extract the email
        return {
            email: row[this.emailColumnIndex],  // Use the provided index for email column
            name: null,     // No names available in this structure (can be customized)
            otherData: []   // No additional data in this structure
        };
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

    async sendEmailForRow(rowIndex, subject, templatePath) {
        try {
            const recipient = this.getRow(rowIndex);
            const replacements = {
                name: recipient.name || 'User', // Use 'User' as a fallback
                ...recipient.otherData.reduce((acc, value, index) => {
                    acc[`other${index + 1}`] = value;
                    return acc;
                }, {})
            };
            const htmlContent = this.createEmailContent(templatePath, replacements);
            await this.sendEmail(recipient.email, subject, htmlContent);
        } catch (error) {
            console.error(error.message);
        }
    }
}

module.exports = AutoEmailSender;
