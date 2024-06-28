const axios = require('axios');

// variais iniciais
let currentPage = 1;
let totalPages = 1;
let allData = [];

// constante para datas. (Nesse caso é constante por que a data não altera, sempre será 01/08/2023 até a data atual.)
const startDate = '2023-08-01';
const endDate = new Date().toISOString().split('T')[0];

// Função para requisição processar a resposta
async function fetchData(page) {
    try {
        const response = await axios.get(`https://api.gestaoclick.com/pagamentos?page=${page}`, {
            headers: {
                'Content-Type': 'application/json',
                'access-token': 'bf07732c8f55a4a4a4891dca2513f9b9e4514136',
                'secret-access-token': 'c6aa5784b15c110c7dc85a8835f32b98870b10f1'
            },
            params:{
                page: page,
                data_inicio: startDate,
                data_fim: endDate
            }
        });

        const data = response.data;

        // printa a resposta da api
        console.log("Resposta da API:", JSON.stringify(data, null, 2));

        // Verificar se a resposta contém o campo resultados
        if (data.resultados && Array.isArray(data.resultados)) {
            allData.push(...data.resultados);
        } else {
            console.error("Campo 'resultados' não encontrado na resposta da API.");
        }

        // Atualizar totalde paginas
        totalPages = data.total_paginas || 1;

        // Verificar se tem mais páginas para processar
        if (currentPage < totalPages) {
            currentPage++;
            await fetchData(currentPage);
        } else {
            // Finalizar e exibir todos os dados
            console.log("Todos os dados coletados:", allData);
        }
    } catch (error) {
        if (error.response) {
            // A resposta foi recebida e o servidor respondeu com um status fora do intervalo 2xx
            console.error("Erro ao fazer a requisição:", error.response.status, error.response.statusText);
            console.error("Detalhes do erro:", error.response.data);
        } else if (error.request) {
            // A requisição foi feita, mas nenhuma resposta foi recebida
            console.error("Nenhuma resposta recebida:", error.request);
        } else {
            // Algo aconteceu ao configurar a requisição que acionou um erro
            console.error("Erro ao configurar a requisição:", error.message);
        }
    }
}

// Iniciar a coleta de dados
fetchData(currentPage);