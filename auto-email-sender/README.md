```markdown
# Auto Email Sender

**Auto Email Sender** is a simple Node.js package that allows you to send emails using an SMTP server. It supports sending emails with customizable HTML content and data extracted from an Excel sheet. This package is ideal for bulk email sending where email content is dynamic and personalized for each recipient.

## Features

- Send emails using any SMTP server.
- Supports HTML templates for email content with dynamic placeholders.
- Extracts recipient information (e.g., email addresses, names, and other data) from an Excel sheet.
- Easy-to-use API with minimal configuration required.

## Installation

Install the package via npm:

```bash
npm install auto-email-sender
```

## Usage

Here’s how you can use the `auto-email-sender` package in your Node.js project:

### Step 1: Prepare Your Excel File

Ensure your Excel file contains the following columns:
- **Email** (the recipient's email address)
- **Name** (the recipient's name)
- **Other columns** (optional data to be included in the email)

Example Excel structure:

| Email               | Name      | Data1 | Data2 |
|---------------------|-----------|-------|-------|
| user1@example.com    | Aniket Subudhi  | Info1 | Info2 |
| user2@example.com    | Swagat| Info3 | Info4 |

### Step 2: Create an HTML Template

Create an HTML file for the email content. Use placeholders in the form of `{{placeholder_name}}` for dynamic content.

```html
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Email Template</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            color: #333;
            margin: 0;
            padding: 20px;
        }
        .container {
            background-color: #fff;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        .header {
            font-size: 24px;
            font-weight: bold;
        }
        .content {
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">Hello {{name}},</div>
        <div class="content">
            This is a dynamic email message.<br>
            Here is some additional data: {{other1}}, {{other2}}.
        </div>
    </div>
</body>
</html>
```

### Step 3: Use the Package in Your Project

Here’s an example of how to use the package:

```javascript
const AutoEmailSender = require('auto-email-sender');

// SMTP configuration
const smtpConfig = {
    host: 'smtp.example.com',
    port: 587,
    secure: false, // true for 465, false for other ports
    auth: {
        user: 'your-email@example.com',
        pass: 'your-email-password',
    },
};

// Specify the column indices (0-based indexing)
const emailColumnIndex = 0; // Column containing email addresses
const nameColumnIndex = 1;  // Column containing names

// Initialize the AutoEmailSender with the Excel file, SMTP config, and column indices
const emailSender = new AutoEmailSender('path/to/your/excel/file.xlsx', smtpConfig, emailColumnIndex, nameColumnIndex);

// Set the subject and path to the HTML template
const subject = 'Your Email Subject';
const templatePath = 'path/to/email_template.html';

// Send an email for a specific row (0-based index)
emailSender.sendEmailForRow(0, subject, templatePath);
```

### Step 4: Run Your Script

To send the emails, simply run your Node.js script:

```bash
node your_script.js
```

### API Documentation

#### `AutoEmailSender(filePath, smtpConfig, emailColumnIndex, nameColumnIndex)`

- **filePath**: Path to the Excel file containing recipient data.
- **smtpConfig**: SMTP configuration object for `nodemailer`. Must include host, port, secure, and auth properties.
- **emailColumnIndex**: 0-based index of the column containing email addresses.
- **nameColumnIndex**: 0-based index of the column containing recipient names.

#### `sendEmailForRow(rowIndex, subject, templatePath)`

- **rowIndex**: The 0-based index of the row in the Excel file to use for the email.
- **subject**: Subject of the email.
- **templatePath**: Path to the HTML template file.

#### Example SMTP Configuration

```javascript
const smtpConfig = {
    host: 'smtp.example.com', // Replace with your SMTP server
    port: 587,
    secure: false, // true for 465, false for other ports
    auth: {
        user: 'your-email@example.com',
        pass: 'your-email-password',
    },
};
```

## License

This project is licensed under the ICT License.

## Author

Aniket Subudhi

## Contributions

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/Aniket-Subudh1/auto-email-sender.git) if you want to contribute.

## Acknowledgments

- [Nodemailer](https://nodemailer.com/) - The Node.js module used for sending emails.
- [xlsx](https://github.com/SheetJS/sheetjs) - The library used to parse Excel files.
- [fs-extra](https://github.com/jprichardson/node-fs-extra) - The library used for file system operations.
```
