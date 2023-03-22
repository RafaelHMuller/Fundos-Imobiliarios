<h1 align="center">
📄<br>README
</h1>

## Índice 

* [Descrição do Projeto](#descrição-do-projeto)
* [Funcionalidades e Demonstração da Aplicação](#funcionalidades-e-demonstração-da-aplicação)
* [Pré requisitos](#pré-requisitos)
* [Execução](#execução)
* [Implantação](#implantação)
* [Bibliotecas](#bibliotecas)

# Descrição do projeto
Projeto Fundos Imobiliários
> Este repositório é meu projeto Python de coleta das principais informações referente à minha carteira de fundos imobiliários. Neste projeto, recebo via e-mail automaticamente valores das cotas e gráficos dos meus fundos de interesse. Criei este projeto não só pelo interesse em acompanhar meus investimentos em fundos imobiliários, mas também para treinar programação, mais especificamente automações na web (web-scrapping) e análise de dados.

# Funcionalidades e Demonstração da Aplicação
Envio de email com informações dos fundos imobiliários:
- gráficos da variação da cotação dos últimos 6 meses
- tabela com o valor atual das cotas
- tabela com a variação das cotas
- tabela com o valor dos dividendos pagos por cota
- valor total atual da carteira
- valor total de dividendos esperado por mês

![Screenshot_3](https://user-images.githubusercontent.com/128300382/227011139-2d5fe5c2-2a66-428d-a16c-56e660fc7f74.png)

## Pré requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Base de dados (arquivo excel)
* Pasta para o armazenamento dos gráficos

## Execução

Neste projeto, há automação (selenium e pyautogui). Durante a automação selenium, através do modo headless, não haverá problema algum ao usuário ao tocar o teclado ou mouse. Entretanto, durante a automação pyautogui, uma mensagem de alerta será mostrada ao usuário, indicando que não use o teclado ou mouse. Ao fim desta automação, outra mensagem de alerta indicará seu final.

## Implantação

É possível adaptar este projeto a qualquer carteira de fundos imobiliários, desde que os códigos dos fundos sejam adicionados à base de dados (arquivo excel).

## Bibliotecas

* selenium
> biblioteca de automação web
* webdriver_manager.chrome
> em conjunto com o selenium, atualiza o drive do Chrome

* pandas
> biblioteca de análise de dados

* win32com.client
> biblioteca que permite a utilização de aplicações do Windows (ex.: Outlook)

* datetime
> biblioteca usada para definir data e horário

* time
> biblioteca que permite definir intervalos de pausa na automação

* PIL
> biblioteca que permite a utilização e edição de imagens

* pyautogui
> biblioteca de automação por meio do mouse, teclado e monitor

* os
> biblioteca de integração de arquivos e pastas do computador
