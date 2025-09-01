# 🏢 Sistema de Inventário Sanofi

Sistema completo para controle de inventário da Sanofi, com interface gráfica e integração SAP.

## 🚀 Funcionalidades

### 📊 Interface Gráfica (Tkinter)
- **Carregamento de Arquivos**: MB52, BMBC, One Portfolio, Inventory
- **Processamento de Dados**: Merge e cálculos automáticos
- **Formatação Excel**: Tabelas e cores automáticas
- **Validações**: Verificação de vencimentos e age limits

### 🔄 Integração SAP
- **Execução MB52**: Relatório de estoque via SAP
- **Execução BMBC**: Relatório de análise via SAP
- **Transferência de Dados**: MB52 → BMBC automática
- **Conversão MHTML**: Para Excel automaticamente

## 🛠️ Instalação

### 1. Instalar Dependências
```bash
pip install -r requirements.txt
```

### 2. Configurar SAP
- SAP GUI instalado
- Acesso ao ambiente PHS
- Permissões para transações MB52 e BMBC

## 🎯 Como Usar

### Interface Gráfica (Recomendado)
```bash
python main_tkinter.py
```

### Linha de Comando
```bash
# Processo completo
python main.py completo

# Relatórios individuais
python main.py mb52
python main.py bmbc
```

## 📋 Fluxo de Trabalho

### 1. **Carregamento de Arquivos**
- Clique nos botões para carregar arquivos Excel
- MB52, BMBC, One Portfolio, Inventory

### 2. **Execução SAP**
- **MB52 (SAP)**: Executa relatório MB52 via SAP
- **BMBC (SAP)**: Executa relatório BMBC via SAP
- **Processo Completo**: MB52 → Extrai Material → BMBC

### 3. **Processamento**
- **Calcular Inventário**: Processa todos os arquivos carregados
- Merge automático entre arquivos
- Formatação e validações

### 4. **Resultado**
- Arquivo Excel formatado: "Controle de Inventário.xlsx"
- Tabelas com cores e validações
- Status de vencimentos e age limits

## 📁 Estrutura de Arquivos

```
Inventário/
├── main_tkinter.py          # Interface gráfica principal
├── main.py                  # Script linha de comando
├── modulo/
│   ├── sap.py              # Integração SAP
│   └── exemplo_uso_sap.py  # Exemplos de uso
├── utils.py                 # Funções de processamento
├── requirements.txt         # Dependências
└── README.md               # Documentação
```

## 🔧 Configurações

### SAP
- Ambiente: PHS [sapphsas01.pharma.aventis.com]
- Caminho SAP: Detectado automaticamente
- Transações: MB52, BMBC

### Arquivos
- **MB52**: Relatório de estoque
- **BMBC**: Relatório de análise
- **One Portfolio**: Dados de age limits
- **Inventory**: Dados de inventário

## 📊 Funcionalidades Avançadas

### Transferência de Dados MB52 → BMBC
1. Executa MB52 via SAP
2. Converte MHTML para Excel
3. Extrai coluna Material
4. Copia para área de transferência
5. Executa BMBC via SAP
6. Cola dados automaticamente

### Processamento de Dados
- Merge automático entre arquivos
- Validação de vencimentos
- Cálculo de age limits
- Remoção de duplicatas
- Formatação de datas

### Formatação Excel
- Tabelas automáticas
- Cores por status
- Cabeçalhos formatados
- Validações visuais

## ⚠️ Observações

1. **SAP GUI**: Deve estar instalado e acessível
2. **Permissões**: Verificar acesso às transações
3. **Arquivos**: Manter estrutura de pastas
4. **Dependências**: Instalar todas as bibliotecas

## 🔍 Troubleshooting

### Erro de Conexão SAP
```
SAP Gui (saplogon.exe) não encontrado
```
**Solução**: Verificar instalação do SAP GUI

### Erro de Permissão
```
Erro ao executar transação
```
**Solução**: Verificar permissões no SAP

### Arquivo não encontrado
```
Arquivo MB52.MHTML não encontrado
```
**Solução**: Executar MB52 primeiro

## 📞 Suporte

Para dúvidas ou problemas:
- Verificar logs de erro
- Consultar documentação SAP
- Verificar permissões de usuário

---

**Desenvolvido para Sanofi - Sistema de Inventário** 🏢 