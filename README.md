<!-- Logo -->
<p align="center">
  <!-- <img src="assets/logo.png" alt="AutoMessenger WhatsApp" width="200"> -->
  <h1 align="center">AutoMessenger WhatsApp</h1>
</p>

<p align="center">
  <strong>Automacao de envio de mensagens via WhatsApp Web</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/Selenium-4.x-43B02A?style=for-the-badge&logo=selenium&logoColor=white" alt="Selenium">
  <img src="https://img.shields.io/badge/WhatsApp-Web-25D366?style=for-the-badge&logo=whatsapp&logoColor=white" alt="WhatsApp">
  <img src="https://img.shields.io/badge/Interface-CustomTkinter-2B2B2B?style=for-the-badge" alt="CustomTkinter">
  <img src="https://img.shields.io/badge/Status-Em%20Desenvolvimento-yellow?style=for-the-badge" alt="Status">
</p>

---

## Sobre

O **AutoMessenger WhatsApp** e uma ferramenta de automacao para envio em massa de mensagens personalizadas via WhatsApp Web. Desenvolvido para otimizar processos de comunicacao empresarial, suportando multiplos modelos de mensagem e estruturas de dados via Excel.

## Funcionalidades

- Envio automatizado de mensagens via WhatsApp Web
- Suporte a multiplos modelos de mensagem (ALL, ONE, ALL_info, Cobranca, ComuniCertificado)
- Importacao de contatos e dados via planilhas Excel (.xlsx)
- Agrupamento automatico por numero de telefone
- Agendamento de envios com keep-alive automatico
- Interface grafica moderna com CustomTkinter
- Sistema de logs detalhado por sessao
- Suporte a multiplos perfis do Chrome
- Formatacao automatica de numeros de telefone (+55)

## Modelos Suportados

| Modelo | Colunas do Excel |
|---|---|
| **ALL** | `Codigo` - `Empresa` - `Contato Onvio` - `Grupo Onvio` - `CNPJ` - `Telefone` |
| **ONE** | `Codigo` - `Nome` - `Numero` - `Caminho` |
| **ALL_info** | `Codigo` - `Nome` - `Numero` + opcionais: `CNPJ`, `Competencia`, `Info_Extra` |
| **Cobranca** | `Codigo` - `Nome` - `Numero` - `Valor da Parcela` - `Data de Vencimento` - `Carta de Aviso` |
| **ComuniCertificado** | `Codigo` - `Nome` - `Numero` - `CNPJ` - `Vencimento` - `Carta de Aviso` |

## Requisitos

```
Python >= 3.10
selenium
webdriver-manager
openpyxl
customtkinter
psutil
Pillow
```

## Instalacao

```bash
# Clonar o repositorio
git clone https://github.com/seu-usuario/AutoMessenger_Whatsa.git
cd AutoMessenger_Whatsa

# Instalar dependencias
pip install selenium webdriver-manager openpyxl customtkinter psutil Pillow
```

## Uso

```bash
python AM_Whatsa.py
```

1. Selecione o **modelo** de mensagem desejado
2. Escolha o **perfil** do Chrome (1, 2 ou Teste)
3. Clique em **Chrome Automacao** para abrir o navegador
4. Carregue o **arquivo Excel** com os contatos
5. Configure a **mensagem** (ou use a mensagem padrao do modelo)
6. Clique em **Iniciar** para comecar o envio

## Estrutura do Projeto

```
AutoMessenger_Whatsa/
|-- AM_Whatsa.py              # Aplicacao principal
|-- mensagens.json            # Templates de mensagens por modelo
|-- AutoMessengerWhatsa_Logs/ # Logs de execucao (ignorado pelo git)
|-- README.md
```

## Perfis do Chrome

| Perfil | Descricao |
|---|---|
| **1** | Perfil de automacao dedicado (`C:\PerfisChrome\automacao_perfil1`) |
| **2** | Perfil de automacao secundario (`C:\PerfisChrome\automacao_perfil2`) |
| **Teste** | Perfil padrao do usuario (copia sessao para perfil isolado) |

---

<p align="center">
  Desenvolvido para automacao de processos internos
</p>
