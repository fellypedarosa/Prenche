# Pr&nche

**Pr&nche** é uma ferramenta prática que preenche automaticamente um modelo de documento Word com dados de uma planilha Excel.

## Descrição

O Pr&nche substitui os espaços reservados (placeholders) no documento modelo pelos valores corretos da planilha, facilitando a criação de documentos padronizados de forma rápida e simples.

## Requisitos

- **Prenche.exe**: O executável do programa (gerado com PyInstaller).
- **Dados de Preenchimento.xlsx**: Planilha com os dados para preencher o modelo.
- **MODELO - Oposição Administrativa - SKO Oyarzabal.docx**: Documento modelo que contém os placeholders a serem substituídos.

> **Importante:** Todos esses arquivos devem estar na mesma pasta. Se algum arquivo não for encontrado, o programa solicitará que você o selecione manualmente.

## Como Usar

1. **Organize os Arquivos:**  
   Coloque os seguintes arquivos na mesma pasta:
   - Prenche.exe
   - Dados de Preenchimento.xlsx
   - MODELO - Oposição Administrativa - SKO Oyarzabal.docx

2. **Execute o Programa:**  
   Dê um duplo clique em `Prenche.exe` para iniciar o programa.

3. **Selecione os Arquivos (se necessário):**  
   Caso algum arquivo não seja encontrado automaticamente, o programa abrirá uma janela para que você os selecione.

4. **Salve o Novo Documento:**  
   O programa processará os dados e, ao final, pedirá para você escolher o local e o nome para salvar o documento preenchido. Um nome padrão será sugerido, mas você pode alterá-lo se desejar.

5. **Revisão Final:**  
   Abra o documento gerado e verifique se todos os campos foram atualizados corretamente. Essa revisão garante que o resultado final esteja perfeito!

## Instalação e Compilação (Opcional)

Se você deseja compilar o programa a partir do código-fonte:

1. Clone este repositório:
   ```bash
   git clone https://github.com/fellypedarosa/Prenche.git
