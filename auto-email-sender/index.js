const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const fs = require('fs-extra');

class AutoEmailSender {
    constructor(filePath, smtpConfig, emailColumnIndex, nameColumnIndex) {
        this.filePath = filePath;
        this.smtpConfig = smtpConfig;
        this.emailColumnIndex = emailColumnIndex;  // Index of the column containing the email addresses
        this.nameColumnIndex = nameColumnIndex;    // Index of the column containing the names
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

        // Use the specified column index to extract the email and name
        return {
            email: row[this.emailColumnIndex],  // Extract email using the provided index
            name: row[this.nameColumnIndex],    // Extract name using the provided index
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
                name: recipient.name || 'User',  // Use the name from the specified column, fallback to 'User'
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
