const { parse } = require('./parse.js');
const { generateSheets } = require('./generate.js');
const { sendEmails } = require('./email.js');

async function main() {
    const data = await parse();
    await generateSheets(data);
    await sendEmails(data);
}

main();