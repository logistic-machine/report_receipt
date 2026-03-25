# Automação de Relatórios (Python)

Este projeto contém ferramentas desenvolvidas para otimizar fluxos de trabalho no setor de logística, integrando Excel, SAP e WMS.

## 🚀 Funcionalidades

### 1. Cruzamento de Bases (WMS vs SAP)
- **Problema:** Dificuldade em conciliar o que foi faturado (Export) com o que chegou fisicamente (WMS).
- **Solução:** Script Python que utiliza a biblioteca `Pandas` para realizar um merge (VLOOKUP) detalhado por item.
- **Diferencial:** Implementação de lógica para evitar duplicidade de valores financeiros em notas com múltiplos produtos.

## 🛠️ Tecnologias Utilizadas
- **Python 3.x**
- **Pandas:** Manipulação de grandes volumes de dados.
- **Openpyxl:** Formatação de planilhas profissionais.

## 📈 Resultados
- Redução de 80% no tempo gasto em tarefas manuais repetitivas.
- Eliminação de erros de digitação e falhas de envio.
