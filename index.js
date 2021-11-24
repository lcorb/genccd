const prompt = require('prompt-sync')();
const { parse } = require('./parse.js');
const { generateSheets } = require('./generate.js');
const { sendEmails } = require('./email.js');

async function main() {
    const data = await parse();
    await generateSheets(data);
    let sendToSelf = true;
    console.log('\nDo you wish to send these emails to yourself to test? (limited to 10)\n(Answering N will send all emails to teachers!)')
    const answer = prompt('\t(Y/N): ');
    console.log('\n')
    if (answer.toLowerCase() === 'y') {
        console.log('Sending emails to self...');
    } else if (answer.toLowerCase() === 'n') {
        console.log('Sending emails to teachers!');
        sendToSelf = false;
    } else {
        console.log('Unrecognised answer, sending emails to self as a precaution...');
    }
    await sendEmails(data, sendToSelf);
}

main();