import pkg from 'portaleargo-api';
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
const { Client } = pkg;

// Check if the required command-line arguments are provided
if (process.argv.length !== 5) {
    console.error('Usage: node MEDIA.mjs CODSCUOLA USERNAME PASSWORD');
    process.exit(1);
}

const schoolCode = process.argv[2];
const username = process.argv[3];
const password = process.argv[4];

const client = new Client({
    schoolCode,
    username,
    password,
});

async function exportToExcel() {
    console.log('elimino cartella temporanea...');
    await fs.rmdir('.argo', { recursive: true });
    console.log('fatto! effettuo login...');

    await client.login();
    var userProfile = await client.getDettagliProfilo();
    console.log('ciao, ' + userProfile.alunno.nome.toLowerCase() + '!');

    await fs.rm('voti-' + userProfile.alunno.nome.toLowerCase() + '.xlsx', { recursive: true });

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

    await workbook.xlsx.writeFile('voti-' + userProfile.alunno.nome.toLowerCase() + '.xlsx');
    console.log('esportato voti!');
}

console.log('ciao! ti ricordo che di solito se ci sono più figli su un account, viene selezionato il primogenito.');
setTimeout(exportToExcel, 1000);


