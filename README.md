# 🍔 Gestor Delivery & KDS (V1.0)

Sistema completo de ERP para Delivery e Restaurantes rodando 100% no Google Apps Script (Serverless).

## 🚀 Funcionalidades
- **App do Cliente:** Cardápio digital via WebApp com envio de pedidos para WhatsApp.
- **KDS (Kitchen Display System):** Monitor de cozinha em tempo real com gestão de status.
- **Logística:** Gestão de Entregadores (Motoqueiros) e taxas por bairro.
- **PDV Balcão:** Módulo para lançamento de pedidos presenciais.
- **Financeiro:** Dashboard automático com DRE e indicadores de venda.
- **Estoque:** Baixa automática via Ficha Técnica.

## 🛠️ Tecnologias
- **Backend:** Google Apps Script (JavaScript Cloud).
- **Database:** Google Sheets (Planilha como Banco de Dados).
- **Frontend:** HTML5, CSS3 e TailwindCSS.

## 📦 Como usar
Este projeto utiliza `clasp` para deploy.
1. Clone o repositório.
2. Instale as dependências: `npm install -g @google/clasp`
3. Faça login: `clasp login`
4. Envie para sua planilha: `clasp push`