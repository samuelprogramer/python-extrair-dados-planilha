// Certifique-se de incluir a biblioteca xlsx
// Você pode instalá-la via npm: npm install xlsx
const XLSX = require('xlsx');

function extrairDadosPlanilha(nomeArquivo, nomePlanilha) {
    // Ler o arquivo Excel
    const workbook = XLSX.readFile(nomeArquivo);

    // Selecionar a planilha pelo nome
    const planilha = workbook.Sheets[nomePlanilha];

    // Converter a planilha para um objeto JSON
    const dados = XLSX.utils.sheet_to_json(planilha, { header: 1 });

    // Iterar sobre as linhas do objeto JSON
    for (const linha of dados.slice(1)) {
        // Supondo que a primeira linha contém os cabeçalhos
        const pergunta = linha[0];
        const alternativas = linha.slice(1, -1); // Excluir a última coluna (resposta)

        // Aqui você pode realizar as operações desejadas com os dados
        console.log(`Pergunta: ${pergunta}`);
        console.log(`Alternativas: ${alternativas}`);
    }
}

// Substitua 'planilhaQuestoes.xlsx' pelo nome do seu arquivo Excel
// Substitua 'Planilha1' pelo nome da sua planilha
extrairDadosPlanilha('planilhaQuestoes.xlsx', 'Planilha1');
