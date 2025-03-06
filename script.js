async function buscarOS() {
    const osNumber = document.getElementById('osNumber').value;

    // URL da planilha publicada
    const url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQjHfFEDNbyJxmtlGysOJb3i35Oq217kDCjU8_4pGVpzMFeOA-qtbC2vV2d_4YG_tR9bcaTue2tr39M/pub?output=xlsx';

    try {
        // Buscar o arquivo XLSX
        const response = await fetch(url);
        const arrayBuffer = await response.arrayBuffer();

        // Ler o arquivo XLSX
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const sheetName = workbook.SheetNames[0]; // Pega a primeira planilha
        const sheet = workbook.Sheets[sheetName];

        // Converter a planilha para JSON
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        // Encontrar a OS correspondente
        const headers = data[0]; // Cabeçalhos
        const osRow = data.find(row => row[0] === osNumber); // Busca a OS

        if (osRow) {
            // Preencher a tabela com os dados
            document.getElementById('submissionDate').textContent = osRow[1] || '';
            document.getElementById('dataAbertura').textContent = osRow[2] || '';
            document.getElementById('centroTrabalho').textContent = osRow[3] || '';
            document.getElementById('tipoManutencao').textContent = osRow[4] || '';
            document.getElementById('descricaoServico').textContent = osRow[5] || '';
            document.getElementById('informacoesComplementares').textContent = osRow[6] || '';
            document.getElementById('numeroSolicitacao').textContent = osRow[7] || '';
            document.getElementById('anexarImagem').textContent = osRow[8] || '';
            document.getElementById('solicitante').textContent = osRow[9] || '';
            document.getElementById('setor').textContent = osRow[10] || '';
            document.getElementById('localAvare').textContent = osRow[11] || '';
            document.getElementById('localHolambra').textContent = osRow[12] || '';
            document.getElementById('localItaberaII').textContent = osRow[13] || '';
            document.getElementById('localSaoManuel').textContent = osRow[14] || '';
            document.getElementById('localTakaoka').textContent = osRow[15] || '';
            document.getElementById('localTaquari').textContent = osRow[16] || '';
            document.getElementById('localTaquarituba').textContent = osRow[17] || '';
            document.getElementById('localTaquarivai').textContent = osRow[18] || '';

            // Exibir a tabela
            document.getElementById('osInfo').classList.remove('hidden');
        } else {
            alert('OS não encontrada!');
        }
    } catch (error) {
        console.error('Erro ao buscar dados:', error);
        alert('Erro ao buscar dados. Verifique o console para mais detalhes.');
    }
}