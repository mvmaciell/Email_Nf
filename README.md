O programa ZOCN_EMAIL_NF é um relatório ABAP desenvolvido para gerenciar e-mails relacionados ao recebimento de Notas Fiscais (NFs) de clientes. O programa realiza a seleção e exibição de dados de NFs, incluindo informações dos clientes e detalhes das faturas, e é capaz de verificar e-mails cadastrados para receber essas NFs.

Funcionalidades
Seleção de Dados:

O usuário pode especificar um intervalo de datas de faturamento, número do cliente, número da ordem de venda e número da NF para filtrar os registros.
Caso um e-mail seja informado, o programa verifica se ele está cadastrado para receber NFs.
Busca de Informações:

O programa realiza consultas em várias tabelas, como J_1BNFDOC, J_1BNFLIN, VBRK, KNA1, ADR6 e VBFA, para obter informações relevantes sobre as NFs e os clientes.
Verifica se o e-mail fornecido está associado ao cliente e está marcado para receber NFs.
Exibição de Dados:

Os dados coletados são organizados e exibidos em uma tabela ALV (ABAP List Viewer) com colunas para o número do cliente, nome do cliente, número da NF, data de faturamento, usuário que criou a fatura e o e-mail para recebimento.
Em caso de sucesso, os dados são exibidos; caso contrário, são apresentadas mensagens de erro ou informações de que nenhum cliente foi encontrado.
Estrutura Técnica
Includes:
ZOCN_EMAIL_NF_TOP: Define tipos de dados e variáveis globais.
ZOCN_EMAIL_NF_SRC: Contém a lógica de seleção de dados e a definição de parâmetros de seleção.
ZOCN_EMAIL_NF_FORM: Implementa os formulários (sub-rotinas) usados para a execução do programa, incluindo a lógica para coleta e apresentação dos dados.

O programa assegura que os e-mails associados aos clientes estejam corretamente registrados para receber as NFs, oferecendo uma interface amigável para visualizar essas informações.
