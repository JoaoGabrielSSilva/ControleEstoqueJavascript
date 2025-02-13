document.addEventListener("DOMContentLoaded", () => {
    const productForm = document.getElementById("productForm");
    const productTable = document.getElementById("productTable");
    const fileInput = document.getElementById("fileInput");
    const forceLoadBtn = document.getElementById("forceLoadBtn");
    const exportBtn = document.getElementById("exportBtn");

    let importedData = []; // Armazena os dados importados
    let products = []; // Lista de produtos

    // Função para renderizar os produtos na tabela
    function renderProducts() {
        productTable.innerHTML = "";
        products.forEach((product, index) => {
            const row = document.createElement("tr");
            row.innerHTML = `
                <td>${product.name}</td>
                <td>
                    <button onclick="updateQuantity(${index}, -1)">-</button>
                    ${product.quantity}
                    <button onclick="updateQuantity(${index}, 1)">+</button>
                </td>
                <td>R$ ${product.price.toFixed(2)}</td>
                <td>
                    <button onclick="editProduct(${index})">Editar</button>
                    <button onclick="deleteProduct(${index})">Excluir</button>
                </td>
            `;
            productTable.appendChild(row);
        });
    }

    // Função para processar dados importados
    function processImportedData() {
        if (importedData.length === 0) {
            alert("Nenhuma planilha foi importada ainda!");
            return;
        }
        products = importedData;
        renderProducts();
    }

    // Importação da planilha XLSX
    fileInput?.addEventListener("change", (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            let jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length < 2) {
                alert("Planilha vazia ou formato incorreto!");
                return;
            }

            const headers = jsonData[0].map(header => header.toLowerCase().trim());
            const nameIndex = headers.indexOf("nome");
            const quantityIndex = headers.indexOf("quantidade");
            const priceIndex = headers.indexOf("preço");

            if (nameIndex === -1 || quantityIndex === -1 || priceIndex === -1) {
                alert("A planilha deve conter as colunas: Nome, Quantidade e Preço.");
                return;
            }

            importedData = jsonData.slice(1).map(row => ({
                name: row[nameIndex] || "Produto Desconhecido",
                quantity: parseInt(row[quantityIndex]) || 0,
                price: parseFloat(row[priceIndex]) || 0
            }));

            alert("Planilha importada com sucesso! Clique em 'Carregar Planilha' para exibir os produtos.");
        };
        reader.readAsArrayBuffer(file);
    });

    // Botão para forçar a exibição dos produtos importados
    forceLoadBtn?.addEventListener("click", processImportedData);

    // Adiciona um novo produto
    productForm?.addEventListener("submit", (event) => {
        event.preventDefault();
        
        const name = document.getElementById("name").value.trim();
        const quantity = parseInt(document.getElementById("quantity").value) || 0;
        const price = parseFloat(document.getElementById("price").value) || 0;

        products.push({ name, quantity, price });
        renderProducts();
        productForm.reset();
    });

    // Atualiza a quantidade de um produto
    window.updateQuantity = (index, change) => {
        if (products[index].quantity + change >= 0) {
            products[index].quantity += change;
            renderProducts();
        }
    };

    // Exclui um produto
    window.deleteProduct = (index) => {
        products.splice(index, 1);
        renderProducts();
    };

    // Edita um produto
    window.editProduct = (index) => {
        const product = products[index];
        document.getElementById("name").value = product.name;
        document.getElementById("quantity").value = product.quantity;
        document.getElementById("price").value = product.price;

        products.splice(index, 1);
        renderProducts();
    };

    // Exporta a tabela para XLSX
    exportBtn?.addEventListener("click", () => {
        const data = products.map(p => ({
            Nome: p.name,
            Quantidade: p.quantity,
            Preço: parseFloat(p.price).toFixed(2)
        }));

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Estoque");
        XLSX.writeFile(wb, "estoque.xlsx");
    });
});
