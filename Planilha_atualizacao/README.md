
# Catland - Automação de Planilha de Controle

Automação criada para simplificar atualização de dados de gatos da ONG Catland. A automação interpreta arquivos com mensagens de Whatsapp e atualiza os campos correspondentes na planilha (processo atualmente realizado por voluntários).


## Funcionalidades

- Validação de nome com Cód. Simplesvet
- Todos os campos da planilha podem ser atualizados
- Uma mesma mensagem pode conter mais de uma atualização
- Log é gerado com sucessos e erros de execução
- Mensagens que não são atualizações são ignoradas automaticamente


## Utilização

Uma vez o projeto esteja apropriadamente configurado, esta automação funciona com os seguintes passos:

- No computador, ir até o Whatsapp Web, grupo Catland Back Office
    - Aplicativo Whatsapp não foi testado
- Copiar todas as mensagens a partir da mensagem "Atualizado até aqui" mais recente (CTRL+C)
- Abrir arquivo *input/mensagens.txt*
- Deletar todo o conteúdo, caso haja
- Colar mensagens (CTRL+V)
- Fechar arquivo *input/mensagens.txt*
- Executar src/main.py
    - Um executável será criado no futuro para simplificar o processo
    - Neste momento todas as mensagens serão processadas e as atualizações serão realizadas na planilha
- Atentar aos erros mostrados no log.
    - Estes precisarão ser preenchidos manualmente na planilha

## Exemplos de Mensagens de Atualização

As mensagens de atualização precisam ser identificadas por uma palavra que indique o início e outra que indique o fim. Por padrão, uma mensagem de atualização deve iniciar com a palavra `Atualização` e terminar com a palavra `Fim!`. Em cada linha, deve ser inserido um parâmetro conforme os exemplos abaixo. Os parâmetros `Nome` e `Cód. Simplesvet` são obrigatórios.

Quaisquer outras mensagens que não possuam as palavras `Atualização` e `Fim!` serão ignoradas pela automação.

#### Exemplos de mensagens

Atualização de Status:
```
Atualização
Nome: Dent 
Cód. Simplesvet: 9431
Status: Adotado
Fim!
```


Atualização de Gênero e Nome:
```
Atualização
Nome: Maltesa
Cód. Simplesvet: 9073
Gênero: Macho
Novo nome: Maltês
Fim!
```

Atualização de Status e Observação de saúde:
```
Atualização
Nome: Vitinho
Cód. Simplesvet: 9426
Status: Estrelinha 
Observação saúde: data do óbito: 11/05/2024 Motivo: complicações das alterações neurológicas
Fim!
```

Atualização de Devolução e Motivo:
```
Atualização
Nome: Níquel
Cód. Simplesvet: 8834
Devolução: 25/04/2024
Motivo: ex-responsável informou que o tempo que precisaria demandar para que felino fizesse as necessidades na caixa de areia era incompatível com o tempo que ela tinha disponível
Local: Hospital Externo
Fim!
```

Todas as atualizações possíveis:
```
Atualização
Nome: Abel Neto
Cód. Simplesvet: 2671
Local: Hospital Externo
Status: Estrelinha 
Observação saúde: data do óbito: 11/05/2024 Motivo: complicações das alterações neurológicas
Novo nome: Abel novo nome
Devolução: 25/04/2024
Motivo: ex-responsável informou que o tempo que precisaria demandar para que felino fizesse as necessidades na caixa de areia era incompatível com o tempo que ela tinha disponível
Entrada ONG: 26/04/2024
Saída ONG: 27/04/2024
Carteirinha de Vacinação: Sim
Data nasc.: 28/04/2024
Gênero: Macho
Raça: SRD
Cor: Preto
Castração: 29/04/2024
FIV: Negativo
FELV: Negativo
Data Teste FIV e FELV: 30/04/2024
Data 2ª Dose Vacina: 01/05/2024
Tipo de Vacina: V3
Renovação Vacina: 02/05/2024
Raiva: 03/05/2024
Renovação Raiva: 04/05/2024
História: A mãe foi encontrada prenha, em uma praça. Pouco depois de ser resgatada deu cria a seis filhotes.
Família: Abel Neto, Tino Marcos, Patricia Poeta, Pedro Bial, Caco Barcelos, Baltazar, Sônia Bridi e Alex Escobar
Observação: Obs na coluna Z
Microchip: Não
Interações com outros animais: Não
Interação com Humanos: N/A
Perfil: Assustado
Fim!
```
## Criação de Conta para acesso por automação

Uma conta de serviço é necessária para acessar a planilha por automação de forma segura. Os passos completos podem ser encontrados em https://pygsheets.readthedocs.io/en/stable/authorization.html. Passos simplificados:
- Acessar Google Cloud: https://console.cloud.google.com/apis/dashboard
- Criar ou escolher projeto
- Habilitar Sheet API
- Criar Conta de Serviço
- Criar chave JSON
- Baixar chave JSON
- Compartilhar a planilha com o e-mail da conta de serviço (compartilhar como Editor)
- Salvar arquivo como *client_secret.json* na pasta *config*


## Configuração

No arquivo config/settings.json, diversas configurações podem ser alteradas. As principais são as seguintes:

`messages.trigger_word`: palavra que indica o início de uma Atualização

`messages.ending_word`: palavra que indica o fim de uma Atualização

`google_sheets.id`: id da planilha a ser alterada

`google_sheets.tab`: aba da planilha a ser alterada


## Autores

- [@ivanajimura](https://www.github.com/ivanajimura)
