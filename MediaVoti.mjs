import pkg from 'portaleargo-api';
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import readline from 'readline';

const { Client } = pkg;

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

async function promptUser(question) {
    return new Promise((resolve) => {
        rl.question(question, resolve);
    });
}

async function exportToExcel(schoolCode, username, password) {
    console.log('elimino cartella temporanea...');
    await fs.rmdir('.argo', { recursive: true });
    console.log('fatto! effettuo login...');

    const client = new Client({
        schoolCode,
        username,
        password,
    });

    await client.login();
    var userData = await client.getDettagliProfilo();
    console.log('ciao, ' + userData.alunno.nome.toLowerCase() + '!');
    await fs.rm('voti-' + userData.alunno.nome.toLowerCase() + '.xlsx', { recursive: true });

    var votiLista = client.dashboard.voti;

    var workbook = new ExcelJS.Workbook();
    var worksheet = workbook.addWorksheet('Voti');

    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };

    worksheet.columns = [
        { header: 'Materia', key: 'materia', width: 20 },
        { header: 'Voto', key: 'voto', width: 10 },
        { header: 'Descrizione', key: 'descrizione', width: 60 }
    ];

    votiLista.forEach((votoEntry, index) => {
        if (votoEntry.numMedia == 1) {
            var materia = votoEntry.desMateria;
            var voto = votoEntry.valore;
            var descrizioneVoto = (votoEntry.descrizioneProva != "" ? votoEntry.descrizioneProva : "Nessuna descrizione");

            var row = worksheet.addRow({ materia, voto, descrizione: descrizioneVoto });

            row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: (index % 2 === 0) ? 'FFFFFF' : 'E6E6E6' } };

            row.font = { bold: true, italic: true };
        }
    });

    var formula = `ROUND(AVERAGE(B1:B${worksheet.rowCount - 1}), 3)`;
    
    var mediaRow = worksheet.addRow({ materia: 'MEDIA', voto: { formula }, descrizione: '' });
    mediaRow.font = { bold: true };
    mediaRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };

    await workbook.xlsx.writeFile('voti-' + userData.alunno.nome.toLowerCase() + '.xlsx');
    console.log('esportato voti!');
}

async function main() {
    const schoolCode = await promptUser('Inserisci il codice scuola: ');
    const username = await promptUser('Inserisci l\'username: ');
    const password = await promptUser('Inserisci la password: ');

    console.log('ciao! Ti ricordo che di solito se ci sono più figli su un account, viene selezionato il primogenito.');
    setTimeout(() => exportToExcel(schoolCode, username, password), 1000);
}

main();

