AUTOMATIZAÇÃO DOS CONTRATOS DO COMERCIAL (SOMENTE DOS TERMINAIS PNG E SJP)

O scrip é simples de ser entendido, é apenas para as meninas do comercial não ficarem indo manualmente ver as datas de vencimento dos contratos delas.
A lógica é a seguinte:  ele conecta no nosso banco SQL diretamente no Escalasoft, filtra os contratos das filiais primeiro (código 1 e 4 no script)
que vencem daqui 30 dias e ai manda um email de alerta nos emails configurados, mas se não tiver nada vencendo nesse dia ele só registra no 'regsitro_auto.txt' e encerra.

Pra rodar na máquina precisa ter o Python instalado e usar o seguinte comando 'pip install pyodbc python-dotenv', para baixar as bibliotecas de banco e outras coisas do ambiente
que usei para configurar.

O arquivo .env tem todos os acessos e senhas, tantos pessoais quanto do banco de dados // EXTREMO CUIDADO PARA NÃO EXIBIR ESSA PASTA FORA DA EMPRESA //. 
Se precisar que seja outros acessos deve ser trocado apenas naquela pasta e mais em nenhum lugar, se não o código quebra.

Sobre a execução, eu deixei configurado no Agendados de Tarefas do windows mesmo para rodar sozinho todo dia as 08hr da manhã. Se precisar testar na mão ou forçar é só abrir
o terminal na pasta e rodar esse código 'py main.py' 

Como eu deixei ele rodando em segundo plano, ele cria um arquivo de registro de log na pasta. Ali fica o histórico de tudo, se rodou certo ou não, erros de banco, se achou 
contratos ou não. Então se algo deu errado é muito RECOMENDADO olhar na pasta de logs primeiro.