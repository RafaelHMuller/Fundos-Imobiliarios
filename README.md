<h1 align="center">
📄<br>README
</h1>

## Projeto Fundos Imobiliários
> Este repositório é meu projeto Python de coleta das principais informações referente à minha carteira de fundos imobiliários. Neste projeto, recebo via e-mail automaticamente valores das cotas e gráficos dos meus fundos de interesse. Criei este projeto não só pelo interesse em acompanhar meus investimentos em fundos imobiliários, mas também para treinar programação, mais especificamente automações na web (web-scrapping) e análise de dados.

## ⚙️ Pré-requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Base de dados (arquivo excel)
* Pasta para o armazenamento dos gráficos

## ⚙️ Executando os testes

Neste projeto, há automação (selenium e pyautogui). Durante a automação selenium, através do modo headless, não haverá problema algum ao usuário ao tocar o teclado ou mouse. Entretanto, durante a automação pyautogui, uma mensagem de alerta será mostrada ao usuário, indicando que não use o teclado ou mouse. Ao fim desta automação, outra mensagem de alerta indicará seu final.

## 📦 Implantação

É possível adaptar este projeto a qualquer carteira de fundos imobiliários, desde que os códigos dos fundos sejam adicionados à base de dados (arquivo excel).

## 🛠️ Construído com

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

## 📄 Licença

Este projeto não possui licença. Disponibilizo para qualquer pessoa interessada.
