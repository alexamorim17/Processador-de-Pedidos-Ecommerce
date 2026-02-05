# Tratamento de Planilha de E-commerce

Este projeto tem como objetivo **tratar e organizar dados exportados da plataforma de e-commerce**.  
O arquivo original (`Report.xlsx`) é obtido via exportação da plataforma e contém diversas informações sobre pedidos, clientes, pagamentos e entregas.  

Ao executar o código Python, o programa irá:  
1. **Carregar o arquivo `Report.xlsx`** utilizando a biblioteca **pandas**.  
2. **Processar e tratar os dados**:  
   - Normalizar colunas e valores.  
   - Filtrar informações relevantes.  
   - Ajustar formatos de datas, valores monetários e documentos.  
   - Criar novas colunas derivadas quando necessário.  
3. **Gerar uma nova planilha Excel (`planilha_nova.xlsx`)** com os dados tratados e organizados para análise ou integração com outros sistemas.  

---

## Tecnologias utilizadas
- **Python 3.x**
- **pandas**: para manipulação e tratamento dos dados.  
- **tkinter**: para criação de uma interface gráfica simples que permite selecionar o arquivo de entrada e salvar o arquivo de saída.  
- **funções personalizadas**: para modularizar o processo de limpeza e transformação dos dados.  

---

## Como executar
1. Certifique-se de ter o **Python 3.x** instalado.  
2. Instale as dependências necessárias:
   ```bash
   pip install pandas openpyxl
3. Execute o script principal:
python main.py

4. Uma janela será aberta (via tkinter) para que você selecione o arquivo Report.xlsx.

5. O programa irá processar os dados e salvar automaticamente uma nova planilha chamada planilha_nova.xlsx no mesmo diretório.

Estrutura esperada

- Entrada: Report.xlsx (arquivo exportado da plataforma de e-commerce).
- Saída: planilha_nova.xlsx (arquivo tratado e pronto para uso).

Exemplo de uso- Arquivo original: contém colunas como Client Name, Payment System Name, SKU Name, Total Value, Status.

- Arquivo tratado: organiza os dados em formato limpo, padronizado e pronto para relatórios ou dashboards.
- Observações- Este projeto pode ser adaptado para diferentes modelos de exportação da plataforma.

- Novas regras de tratamento podem ser adicionadas conforme a necessidade do negócio.
