import pkg from 'portaleargo-api';
import ExcelJS from 'exceljs';
import fs from 'fs/promises';  // Import the fs module for file system operations
const { Client } = pkg;

const client = new Client({
    schoolCode: "SCUOLACODE",
    username: "USERNAME",
    password: "PASSWORD",
});

async function exportToExcel() {
    // Eliminate the ".argo" folder at each execution
    console.log('elimino cartella temporanea...');
    await fs.rmdir('.argo', { recursive: true });
    console.log('fatto! effettuo login...');

    await client.login();
    var username = await client.getDettagliProfilo();
    console.log('ciao, ' + username.alunno.nome.toLowerCase() + '!');
    await fs.rm('voti-' + username.alunno.nome.toLowerCase() + '.xlsx', { recursive: true });

    var votiLista = client.dashboard.voti;

    // Create a new workbook and worksheet with styling
    var workbook = new ExcelJS.Workbook();
    var worksheet = workbook.addWorksheet('Voti');
    
    // Styling for header row
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } }; // Yellow background

    // Styling for data rows
    worksheet.columns = [
        { header: 'Materia', key: 'materia', width: 20 },
        { header: 'Voto', key: 'voto', width: 10 },
        { header: 'Descrizione', key: 'descrizione', width: 60 }
    ];

    votiLista.forEach((votoEntry, index) => {
        if (votoEntry.numMedia == 1)
        {
            var materia = votoEntry.desMateria;
            var voto = votoEntry.valore;
            var descrizioneVoto = (votoEntry.descrizioneProva != "" ? votoEntry.descrizioneProva : "Nessuna descrizione");

            var row = worksheet.addRow({ materia, voto, descrizione: descrizioneVoto });

            // Alternate row background color
            row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: (index % 2 === 0) ? 'FFFFFF' : 'E6E6E6' } };

            // Bold and italicize text
            row.font = { bold: true, italic: true };
        }
    });

    // Calculate the average of all grades in the second column (Voto)
    var formula = `ROUND(AVERAGE(B1:B${worksheet.rowCount - 1}), 3)`; // Use ROUND function to round the average
    
    // Styling for the "MEDIA" row
    var mediaRow = worksheet.addRow({ materia: 'MEDIA', voto: { formula }, descrizione: '' });
    mediaRow.font = { bold: true };
    mediaRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } }; // Yellow background

    // Save the workbook to a file
    await workbook.xlsx.writeFile('voti-' + username.alunno.nome.toLowerCase() + '.xlsx');
    console.log('esportato voti!');
}

console.log('ciao! ti ricordo che di solito se ci sono più figli su un account, viene selezionato il primogenito.');
setTimeout(exportToExcel, 1000);
