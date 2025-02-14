document.addEventListener("DOMContentLoaded", () => {
    // Obtenção dos elementos do DOM
    const formularioProduto = document.getElementById("formularioProduto");  // Formulário para adicionar produtos
    const tabelaProdutos = document.getElementById("tabelaProdutos");    // Tabela onde os produtos serão exibidos
    const inputArquivo = document.getElementById("inputArquivo");        // Entrada de arquivo para importar a planilha
    const botaoForcarCarregar = document.getElementById("botaoForcarCarregar"); // Botão para carregar os produtos manualmente
    const botaoExportar = document.getElementById("botaoExportar");        // Botão para exportar os dados da tabela

    let dadosImportados = []; // Armazena os dados importados da planilha
    let produtos = [];        // Lista de produtos

    // Função para renderizar os produtos na tabela
    function renderizarProdutos() {
        tabelaProdutos.innerHTML = ""; // Limpa a tabela
        produtos.forEach((produto, indice) => {
            const linha = document.createElement("tr");
            linha.innerHTML = `
                <td>${produto.nome}</td>
                <td>
                    <button onclick="atualizarQuantidade(${indice}, -1)">-</button>
                    ${produto.quantidade}
                    <button onclick="atualizarQuantidade(${indice}, 1)">+</button>
                </td>
                <td>R$ ${produto.preco.toFixed(2)}</td>
                <td>
                    <!-- Exibe o tamanho diretamente -->
                    ${produto.tamanho}
                </td>
                <td>
                    <button onclick="editarProduto(${indice})">Editar</button>
                    <button onclick="excluirProduto(${indice})">Excluir</button>
                </td>
            `;

            tabelaProdutos.appendChild(linha);
        });
    }

    // Função para processar os dados importados
    function processarDadosImportados() {
        if (dadosImportados.length === 0) {
            alert("Nenhuma planilha foi importada ainda!");
            return;
        }
        produtos = dadosImportados; // Atualiza a lista de produtos com os dados importados
        renderizarProdutos();       // Chama a função para renderizar os produtos na tabela
    }

    // Importação da planilha XLSX
    inputArquivo?.addEventListener("change", (evento) => {
        const arquivo = evento.target.files[0];  // Obtém o arquivo selecionado
        if (!arquivo) return;  // Se não houver arquivo, sai da função

        const leitor = new FileReader();  // Cria o objeto FileReader para ler o arquivo
        leitor.onload = (e) => {
            const dados = new Uint8Array(e.target.result);  // Obtém o conteúdo do arquivo como array de bytes
            const livro = XLSX.read(dados, { type: "array" });  // Lê os dados da planilha XLSX
            const nomeAba = livro.SheetNames[0];  // Obtém o nome da primeira aba da planilha
            const aba = livro.Sheets[nomeAba];   // Obtém a primeira aba da planilha
            let dadosJson = XLSX.utils.sheet_to_json(aba, { header: 1 });  // Converte a aba para JSON

            if (dadosJson.length < 2) {
                alert("Planilha vazia ou formato incorreto!");
                return;
            }

            const cabecalhos = dadosJson[0].map(cabecalho => cabecalho.toLowerCase().trim());  // Converte os cabeçalhos para minúsculas e remove espaços
            const indiceNome = cabecalhos.indexOf("nome");
            const indiceQuantidade = cabecalhos.indexOf("quantidade");
            const indicePreco = cabecalhos.indexOf("preço");
            const indiceTamanho = cabecalhos.indexOf("tamanho");  // Certifique-se de que há uma coluna de "tamanho"

            if (indiceNome === -1 || indiceQuantidade === -1 || indicePreco === -1 || indiceTamanho === -1) {
                alert("A planilha deve conter as colunas: Nome, Quantidade, Preço e Tamanho.");
                return;
            }

            // Converte os dados da planilha para um formato adequado para os produtos
            dadosImportados = dadosJson.slice(1).map(linha => ({
                nome: linha[indiceNome] || "Produto Desconhecido",  // Define nome padrão caso não haja valor
                quantidade: parseInt(linha[indiceQuantidade]) || 0,   // Define quantidade padrão caso não haja valor
                preco: parseFloat(linha[indicePreco]) || 0,           // Define preço padrão caso não haja valor
                tamanho: linha[indiceTamanho] || "P"                  // Define tamanho padrão caso não haja valor
            }));

            alert("Planilha importada com sucesso! Clique em 'Carregar Planilha' para exibir os produtos.");
        };
        leitor.readAsArrayBuffer(arquivo);  // Lê o arquivo como ArrayBuffer
    });

    // Botão para forçar a exibição dos produtos importados
    botaoForcarCarregar?.addEventListener("click", processarDadosImportados);

    // Adiciona um novo produto
    formularioProduto?.addEventListener("submit", (evento) => {
        evento.preventDefault();  // Previne o comportamento padrão do formulário (envio)

        const nome = document.getElementById("nome").value.trim();     // Obtém o nome do produto
        const quantidade = parseInt(document.getElementById("quantidade").value) || 0;  // Obtém a quantidade do produto
        const preco = parseFloat(document.getElementById("preco").value) || 0;  // Obtém o preço do produto
        const tamanho = document.getElementById("tamanho").value;  // Captura o tamanho selecionado

        produtos.push({ nome, quantidade, preco, tamanho });  // Adiciona o novo produto à lista
        renderizarProdutos();  // Atualiza a tabela com os novos dados
        formularioProduto.reset();  // Limpa o formulário
    });

    // Atualiza o tamanho de um produto
    window.atualizarTamanho = (indice, novoTamanho) => {
        produtos[indice].tamanho = novoTamanho;  // Atualiza o tamanho do produto
        renderizarProdutos();  // Atualiza a tabela
    };

    // Atualiza a quantidade de um produto
    window.atualizarQuantidade = (indice, alteracao) => {
        if (produtos[indice].quantidade + alteracao >= 0) {  // Verifica se a quantidade não pode ficar negativa
            produtos[indice].quantidade += alteracao;
            renderizarProdutos();  // Atualiza a tabela
        }
    };

    // Exclui um produto
    window.excluirProduto = (indice) => {
        produtos.splice(indice, 1);  // Remove o produto da lista
        renderizarProdutos();  // Atualiza a tabela
    };

    // Edita um produto
    window.editarProduto = (indice) => {
        const produto = produtos[indice];
        document.getElementById("nome").value = produto.nome;
        document.getElementById("quantidade").value = produto.quantidade;
        document.getElementById("preco").value = produto.preco;
        document.getElementById("tamanho").value = produto.tamanho;  // Atualiza o valor do tamanho ao editar

        produtos.splice(indice, 1);  // Remove o produto da lista (ele será adicionado novamente após a edição)
        renderizarProdutos();  // Atualiza a tabela
    };

    // Exporta a tabela para XLSX
    botaoExportar?.addEventListener("click", () => {
        const dados = produtos.map(p => ({
            Nome: p.nome,
            Quantidade: p.quantidade,
            Preço: parseFloat(p.preco).toFixed(2),
            Tamanho: p.tamanho
        }));

        const aba = XLSX.utils.json_to_sheet(dados);  // Converte os dados para uma aba de planilha
        const livro = XLSX.utils.book_new();  // Cria um novo livro de trabalho
        XLSX.utils.book_append_sheet(livro, aba, "Estoque");  // Adiciona a aba ao livro de trabalho
        XLSX.writeFile(livro, "estoque.xlsx");  // Exporta o arquivo XLSX
    });
});
