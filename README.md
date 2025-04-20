# Grafico-curvaS

Este projeto foi desenvolvido por mim durante o meu estágio, com o objetivo de calcular semanalmente as atividades planejadas e concluídas. A partir desses dados, é possível gerar um gráfico de linha representando a Curva S de atividades, diretamente no Google Sheets.

O sistema acessa uma pasta que contém as planilhas individuais dos projetos da área de engenharia e cruza essas informações com os dados provenientes de um arquivo do Microsoft Project — que contém o plano mensal da empresa e foi previamente convertido para uma planilha no Google Sheets.

Todos os dados são processados e organizados automaticamente em abas auxiliares de cada planilha, servindo como base para a construção do gráfico, desenvolvido dentro do próprio Google Sheets.

Além disso, o projeto conta com uma função que calcula a semana do ano seguindo o padrão ISO 8601, garantindo precisão na categorização temporal das atividades.