document.getElementById('file1').addEventListener('change', handleFileSelect);
document.getElementById('file2').addEventListener('change', handleFileSelect);
document.getElementById('generateReport').addEventListener('click', generateReport);
document.getElementById('generatePDF').addEventListener('click', generatePDF);

let file1Data, file2Data;

function handleFileSelect(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            if (event.target.id === 'file1') {
                file1Data = workbook;
            } else {
                file2Data = workbook;
            }

            if (file1Data && file2Data) {
                compareWorkbooks(file1Data, file2Data);
            }
        } catch (error) {
            alert('Erro ao ler o arquivo: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function compareWorkbooks(workbook1, workbook2) {
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = 'Comparando arquivos...';

    setTimeout(() => {
        resultDiv.innerHTML = '';

        workbook1.SheetNames.forEach((sheetName, index) => {
            const sheet1 = XLSX.utils.sheet_to_json(workbook1.Sheets[sheetName], { header: 1 });
            const sheet2 = workbook2.SheetNames[index] ? XLSX.utils.sheet_to_json(workbook2.Sheets[workbook2.SheetNames[index]], { header: 1 }) : [];

            const table = document.createElement('table');
            const thead = document.createElement('thead');
            const tbody = document.createElement('tbody');

            const headerRow = document.createElement('tr');
            headerRow.innerHTML = `<th>${sheetName} - Arquivo 1</th><th>Status</th><th>${workbook2.SheetNames[index] || 'Sem correspondência'} - Arquivo 2</th>`;
            thead.appendChild(headerRow);
            table.appendChild(thead);

            const maxRows = Math.max(sheet1.length, sheet2.length);
            const maxCols1 = sheet1[0] ? sheet1[0].length : 0;
            const maxCols2 = sheet2[0] ? sheet2[0].length : 0;
            const maxCols = Math.max(maxCols1, maxCols2);

            for (let i = 0; i < maxRows; i++) {
                const row = document.createElement('tr');
                for (let j = 0; j < maxCols; j++) {
                    const cell1 = (sheet1[i] && sheet1[i][j]) || '';
                    const cell2 = (sheet2[i] && sheet2[i][j]) || '';

                    const td1 = document.createElement('td');
                    const tdStatus = document.createElement('td');
                    const td2 = document.createElement('td');

                    td1.textContent = cell1;
                    td2.textContent = cell2;

                    td1.style.textAlign = 'center';
                    tdStatus.style.textAlign = 'center';
                    td2.style.textAlign = 'center';

                    if (cell1 === cell2) {
                        td1.classList.add('duplicate');
                        td2.classList.add('duplicate');
                        tdStatus.textContent = '0';
                    } else {
                        td1.classList.add('different');
                        td2.classList.add('different');
                        tdStatus.textContent = 'divergência';
                        tdStatus.style.backgroundColor = 'yellow';
                    }

                    row.appendChild(td1);
                    row.appendChild(tdStatus);
                    row.appendChild(td2);
                }
                tbody.appendChild(row);
            }

            table.appendChild(tbody);
            resultDiv.appendChild(table);
        });
    }, 1000);
}

function generateReport() {
    if (!file1Data || !file2Data) {
        alert('Por favor, selecione os dois arquivos Excel.');
        return;
    }

    const workbook = XLSX.utils.book_new();

    file1Data.SheetNames.forEach((sheetName, index) => {
        const sheet1 = XLSX.utils.sheet_to_json(file1Data.Sheets[sheetName], { header: 1 });
        const sheet2 = file2Data.SheetNames[index] ? XLSX.utils.sheet_to_json(file2Data.Sheets[file2Data.SheetNames[index]], { header: 1 }) : [];

        const maxRows = Math.max(sheet1.length, sheet2.length);
        const maxColsFile1 = sheet1[0] ? sheet1[0].length : 0;
        const maxColsFile2 = sheet2[0] ? sheet2[0].length : 0;

        const sheetData = [];

        // Cabeçalhos
        const headerRow = [];
        if (sheet1[0]) {
            headerRow.push(...sheet1[0]);
        }
        headerRow.push('Status');
        if (sheet2[0]) {
            headerRow.push(...sheet2[0]);
        }
        sheetData.push(headerRow);

        // Adicionar dados
        for (let i = 1; i < maxRows; i++) {
            const row = [];
            for (let j = 0; j < maxColsFile1; j++) {
                row.push(sheet1[i] ? sheet1[i][j] : '');
            }
            if (sheet1[i] && sheet2[i] && sheet1[i].join() === sheet2[i].join()) {
                row.push('0');
            } else {
                row.push('divergência');
            }
            for (let j = 0; j < maxColsFile2; j++) {
                row.push(sheet2[i] ? sheet2[i][j] : '');
            }
            sheetData.push(row);
        }

        // Remover colunas em branco
        const nonEmptyCols = [];
        for (let col = 0; col < sheetData[0].length; col++) {
            if (sheetData.some(row => row[col] !== '')) {
                nonEmptyCols.push(col);
            }
        }

        const filteredSheetData = sheetData.map(row => nonEmptyCols.map(col => row[col]));

        const newSheet = XLSX.utils.aoa_to_sheet(filteredSheetData);
        XLSX.utils.book_append_sheet(workbook, newSheet, sheetName);
    });

    XLSX.writeFile(workbook, 'Relatorio_Comparacao_Lado_a_Lado.xlsx');
}
