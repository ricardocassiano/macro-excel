# macro-excel
Planilha com macros programadas em VB

Para ativar o VBA no Excel, após abrir a planilha, basta utilizar o comando ALT+F11
Para vincular uma macro a um botão no Excel, basta ativar a Guia Desenvolvedor e inserir um botão Controle De Formulário. Ao inserir o botão, vincular a macro programada.
As macros são funções criadas com o comando Sub. Numa planilha sem macros, é necessário criar um módulo (Inserir>Módulo - no VBA)

A planilha do Excel possui duas funções macros programadas:

1 - Inserir dados automaticamente ao clicar no botão Inserir Dados:
  Esta função verifica qual é a última linha da tabela formulário preenchida, define o intervalo (Range) de inserção dos dados e adiciona as linhas conforme a quantidade digitada na tabela usada como formulário.
2 - Limpar dados da planilha ao clicar no botão Limpar dados:
  Limpa os dados de determinado intervalo declarado.
