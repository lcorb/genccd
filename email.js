const prompt = require('prompt-sync')();
const EQ_USER = prompt('Enter your email: ');
const EQ_PASS = prompt('Enter your password: ', {echo: '*'});
const nodemailer = require('nodemailer');
const { format } = require('./email_template');

async function sendEmails(data, sendToSelf=true) {
    let emails = [];
    for (const teacher in data) {
        emails.push(formatEmail(teacher, data[teacher]));
    }

    emails = emails.slice(0, 9);

    await queueEmails(emails, 'lcorb28@eq.edu.au', EQ_USER, EQ_PASS, sendToSelf);
}

function formatEmail(teacher, data) {
    let students = {}
    let keys = Object.values(data.classes);
    keys.forEach(c => {
        if (c !== undefined) {
            Object.values(c).forEach(s => { students[s] = true });
        }
    });

    let studentCount = Object.values(students).length;

    return {
        message: format(`Hi ${data.f_name}`, 'Your NCCD data for the year is ready to be completed. Please download and complete the attached file, and send it back to <a href="mailto:dkeen27@eq.edu.au?subject=Completed NCCD Spreadsheet">David Keenan</a> with the subject "Completed NCCD Spreadsheet".', `${studentCount} student${studentCount > 1 ? 's' : ''}`),
        subject: 'NCCD Data',
        to: data.email,
        attachment: `./sheets/${data.f_name} ${data.l_name} - NCCD.xlsx`
    };
}

async function queueEmails(emails, from, user, pass, sendToSelf=true) {
    let transport = nodemailer.createTransport({
        host: "smtp.office365.com", // hostname smtp-mail.outlook.com
        secureConnection: false, // TLS requires secureConnection to be false
        port: 587, // port for secure SMTP
        auth: {
            user: user,
            pass: pass
        },
        tls: {
            ciphers:'SSLv3'
        }
    });

    for (const email of emails) {
        const m = {
          from: from,
          to: sendToSelf ? user: email.to,
          subject: email.subject,
          html: email.message,
          attachments: [{
              path: email.attachment
          }]
        };
        await transport.sendMail(m);
    }

    console.log(`Sent ${Object.keys(emails).length} emails.`);
}

module.exports = { sendEmails }