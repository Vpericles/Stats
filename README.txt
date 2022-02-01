READ ME
__________________________
STATS (SARIMA model)
Vitor Péricles de Carvalho
vitor@laic.com.br
__________________________

Propósito:
- Código simples para ajuste de um modelo SARIMA a uma série ou conjunto de séries temporais, e um forecast para o próximo dado a partir do modelo.

Inputs:
- Arquivo .XLS com nome(s) da(s) série(s) e série(s) histórica(s) de dados;

Outputs:
- calcula um forecast (próximo dado) para a série ou séries de dados a partir do modelo SARIMA;
- salva arquivo .XLS com este(s) forecast(s);

Bibliotecas utilizadas:
- statsmodels (para criação do modelo SARIMAX);
- xlrd (lê arquivo de Excel);
- xlsxwriter (grava arquivo de Excel);
- pandas;
- numpy;
- math;
- matplotlib (para pyplot, caso queira plotar algo);

__________________________

O código está programado para aderência de um modelo SARIMA a uma ou mais séries de dados em um arquivo de Excel.
Após a aderência do modelo, ele também cria um forecast (próximo dado), chamado de Next Step, para servir como
projeção de um "próximo valor" para a série ou conjunto de séries, e salva este(s) resultado(s) em um Excel.
