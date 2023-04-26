
Mudanças de pratica:
    - DATAS: ajustar as datas na planilha (só deixar as datas dos meses usuais) - no segundo calendario de Despesas Indiretas

Tratamentos no PBI:
    - Filtra coluna "VALOR" para -> mudar <type> para texto -> substituir "." para "," -> mudar <type> para valor, float

Arquitetura:
    - os quatro arquivos .txt são arquivos de controle para otimizar o processo

log.txt ->  0 - framework_PBI não gerado (primeira vez) / 
            1 - framework_PBI existe, seguir tratativas

historico.txt -> primeiro historico rodado em log=0 e alerado após toda execução com o valor de historico2

historico2.txt -> historico auxiliar que lê o conteudo atual da pasta e confere divergências com o primeiro historico

framework_PBI.txt -> DataFrame principal salvo em txt para evitar a releitur de arquivo já otimizaados 

IMPORTANTE: 
    - não apagar o arquivo "log.txt"
    - para resetar a quarry e reconferir todos os valores mudar log para 1



PBI_result -> tabela a ser puxada no PowerBI


Funções possíveis mas o JP não tinha tempo : 
    - Enviar um relatório por email de erros ou inconsitência nos valores
    - Caso o erro seja com relação a exclusão de linhas, criar uma função para completar as linhas faltantes (manipulação de arquivo)

