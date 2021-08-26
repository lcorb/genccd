const creds = require('./creds.json');
const EQ_USER = Buffer.from(creds.user, 'base64').toString('binary');
const EQ_PASS = Buffer.from(creds.pw, 'base64').toString('binary');
const nodemailer = require('nodemailer');
const { format } = require('./email_template');

async function sendEmails(data) {
    let emails = [];
    for (const teacher in data) {
        emails.push(formatEmail(teacher, data[teacher]));
    }
    
    await queueEmails([emails[0]], 'lcorb28@eq.edu.au', EQ_USER, EQ_PASS);
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
    // return {
    //     message: `Hi ${data.f_name},\nFor NCCD, you have ${studentCount} students to complete across ${Object.values(data.classes).length} classes.\n\n${data.email}`,
    //     subject: 'NCCD',
    //     to: 'lcorb28@eq.edu.au',
    //     attachment: `./sheets/${data.f_name} ${data.l_name} - NCCD.xlsx`
    // };

    return {
        message: format(`Hi ${data.f_name}`, 'Your NCCD data is ready to be completed. Please download and complete the attached file, and send it back to <a href="mailto:dkeen27@eq.edu.au?subject=Completed NCCD Spreadsheet">David Keenan</a> with the subject "Completed NCCD Spreadsheet".', `${studentCount} student${studentCount > 1 ? 's' : ''}`),
        subject: 'NCCD',
        to: 'lcorb28@eq.edu.au',
        attachment: `./sheets/${data.f_name} ${data.l_name} - NCCD.xlsx`
    };
}

async function queueEmails(emails, from, user, pass) {
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
          to: email.to,
          subject: email.subject,
          html: email.message,
          attachments: [{
              path: email.attachment
          }]
        };
        await transport.sendMail(m);
    }
}

module.exports = { sendEmails }