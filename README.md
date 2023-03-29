<h1 align="center">
📄<br>README - Projeto Automação, Análise de Dados, Dashboard
</h1>

## Índice 

* [Descrição do Projeto](#descrição-do-projeto)
* [Funcionalidades e Demonstração da Aplicação](#funcionalidades-e-demonstração-da-aplicação)
* [Pré requisitos](#pré-requisitos)
* [Execução](#execução)
* [Bibliotecas](#bibliotecas)

# Descrição do projeto
> Este repositório é meu projeto Python de automação local, análise de dados e dashboard online de uma grande rede fictícia de lojas de roupa com 25 lojas espalhadas por todo o Brasil. Inicialmente, o projeto consiste em automatizar diariamente a criação de uma base de dados específica para cada loja a partir das bases de dados geral de toda a rede. Em seguida, são enviados e-mails para cada um dos gerentes das lojas com um One Page informativo dos indicadores financeiros diários de interesse da rede. Posteriormente, a diretoria da rede recebe também um e-mail com o ranking atualizado das melhores lojas do dia e do ano. Por fim, são criados gráficos e estes são acessíveis ao usuário, no caso a diretoria da rede, em um Dashboard interativo online.

# Funcionalidades e Demonstração da Aplicação

E-mail enviado a cada um dos gerentes com um One Page informativo do dia:<br>
![Screenshot_1](https://user-images.githubusercontent.com/128300382/228629549-00fc0d85-4ee2-452d-b703-f7d71414344e.png)
<br>
Parte inicial e final do e-mail enviado à diretoria da rede com os rankings diário e anual:<br>
![Screenshot_2](https://user-images.githubusercontent.com/128300382/228629680-80d0103c-472e-4bec-b7b2-4fb881d42f9c.png)
![Screenshot_3](https://user-images.githubusercontent.com/128300382/228629685-a8239939-0d83-49f9-8eee-b1419cf0b2f0.png)
<br>
Dashboard online:<br>
![Screenshot_4](https://user-images.githubusercontent.com/128300382/228629805-e0aeae4d-4a2d-4348-8c62-1d03d2c42b2c.png)

## Pré requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Base de dados (arquivos Excel e csv na pasta "Base de Dados")
* Pasta com as bases de dados específicas criadas para cada loja (arquivos Excel criados na pasta "Backup Arquivos Lojas - Revisão")
* Navegador web (para o acesso ao Gmail e Dashboard)

## Execução

Ao executar o código, de maneira automática, todo o passo a passo contido na descrição deste Readme será executado e, consequentemente, os e-mails serão enviados para um Gmail. Para alterar o local de envio dos e-mails, deve-se alterar o arquivo "Emails.xlsx" na pasta "Base de Dados". Por fim, para obter acesso ao Dashboard online, é necessário clicar no link enviado junto ao e-mail da diretoria.

## Bibliotecas

* <strong>pandas:</strong> bibliotecas de integração de arquivos excel, csv e outros, possibilitando análise de dados<br>
* <strong>pathlib:</strong> biblioteca de integração de arquivos e pastas do computador<br>
* <strong>win32com.client:</strong> biblioteca de integração dos aplicativos Windows, no caso, do Outlook<br>
* <strong>plotly.express, matplotlib.pyplot, seaborn:</strong> biblioteca de criação de gráficos<br>
* <strong>dash:</strong> biblioteca de criação de Dashboard online<br>
