# Plano de Implementação: Dashboard de Entradas Mensais

Este documento detalha o plano para a criação de uma página dedicada à análise das notas fiscais de entrada (NFs), conforme solicitado.

## 📋 Visão Geral
Criação de uma interface premium para visualizar o volume financeiro recebido no estoque, com métricas de total por período e distribuição por fornecedores.

## 🎯 Objetivos
- Criar a página `entradas.html` com design consistente ao sistema atual.
- Implementar processamento de arquivos Excel (.xlsx) com foco nas colunas: `Dt.movto`, `Razão social` e `Vlr.cont`.
- Gerar cartões (cards) de resumo e gráficos de distribuição por fornecedor.

## 🏗️ Estrutura de Arquivos
- `./entradas.html`: Nova página principal da análise de entradas.
- `./entradas.js`: Lógica de processamento de dados e renderização de gráficos.
- `./styles.css`: Reutilização e expansão dos estilos existentes (Premium Dark Mode).

## 🛠️ Tecnologias
- **HTML5/JS/CSS3** (Vanilla)
- **SheetJS (XLSX)**: Leitura de arquivos Excel.
- **Chart.js**: Visualização de dados.

## 🗓️ Cronograma de Tarefas

### Fase 1: Fundação
- [ ] **T1: Criar `entradas.html`**
  - Estrutura base seguindo o layout do Gestor de Estoque.
  - Seção de cabeçalho com navegação de volta.
  - Área de upload dedicada para NFs.
- [ ] **T2: Adaptar `styles.css`**
  - Garantir que os novos componentes de cards e gráficos da página de entradas mantenham a estética premium.

### Fase 2: Lógica de Dados
- [ ] **T3: Criar `entradas.js`**
  - Implementar `handleFileUpload` específico para as colunas das NFs.
  - Função de processamento: Agrupar por `Razão social` e somar `Vlr.cont`.
  - Formatação de datas a partir de `Dt.movto`.
- [ ] **T4: Implementar Cards de Resumo**
  - Card: **Total Recebido (Mês)**.
  - Card: **Qtd. de Notas Processadas**.
  - Card: **Maior Fornecedor do Mês**.

### Fase 3: Visualização (UI/UX Pro Max)
- [ ] **T5: Gráfico de Fornecedores**
  - Gráfico de barras horizontais (Horizontal Bar) ou Donut para o Top 10 fornecedores por valor.
- [ ] **T6: Tabela de Detalhes**
  - Listagem detalhada das NFs com busca e filtros.

### Fase 4: Integração
- [ ] **T7: Navegação entre Páginas**
  - Adicionar link na `index.html` para acessar o Dashboard de Entradas.
  - Adicionar link de volta na `entradas.html`.

## ✅ Critérios de Sucesso
- [ ] Carregamento de arquivos `.xlsx` sem erros.
- [ ] Visualização clara do total recebido em Reais (R$).
- [ ] Ranking de fornecedores gerado automaticamente.
- [ ] Design responsivo e esteticamente alinhado ao projeto original (Sem usar roxo/púrpura).

---
**Próximo Passo:** Iniciar a criação do arquivo `entradas.html`.
