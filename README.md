Este é um sistema completo de mala direta desenvolvido em Python com Flask, projetado para equipes que precisam enviar campanhas de email personalizadas em massa. O sistema permite que múltiplos usuários trabalhem em conjunto sem conflitos, garantindo que apenas o criador de cada mala possa editá-la, enquanto qualquer membro da equipe pode realizar o envio.

Funcionalidades Principais

Controle de Acesso por Usuário
- Apenas quem criou a mala pode editá-la
- Qualquer membro da equipe pode enviar a mala
- Duplicação de registros: ao duplicar, o novo usuário se torna dono da cópia e pode editá-la livremente
- A mala original permanece intacta com seu criador original

Flexibilidade na Origem dos Dados
- Anexar arquivo: envia o arquivo original junto com o email
- Usar DataFrame: utiliza os dados do arquivo sem anexá-lo
- Ambos: envia tanto os dados quanto o arquivo anexado
- Nenhum: envia apenas o texto puro, sem dados da planilha

Editor de Mensagens 
- Variáveis como {destinatario} (substituído pelo nome do cliente)
- {dataframe} para inserir tabelas com os dados filtrados de cada destinatário
- Cálculos automáticos no corpo do email: soma, média, contagem, valor máximo e mínimo
- Formatação em negrito e cores

Gerenciamento de Campanhas
- Upload de arquivos Excel (.xlsx, .xls) e CSV
- Seleção de planilha e colunas desejadas
- Verificação automática de duplicatas antes de salvar
- Histórico completo com filtros por usuário, data e palavra-chave
- Modo de teste: envia até 5 emails apenas para o próprio usuário
- Reprocessamento de malas antigas com um clique

Tecnologias Utilizadas
- Backend: Python com Flask
- Manipulação de dados: Pandas e OpenPyXL
- Integração com email: Win32com (Outlook)
- Frontend: HTML, CSS, JavaScript puro

Estrutura de Pastas
Projeto_Mala_Direta/
├── app.py                      # Arquivo principal da aplicação
├── templates/                   # Páginas HTML
│   ├── index.html               # Página inicial
│   ├── upload.html              # Criação de nova mala
│   ├── table.html               # Histórico de malas
│   ├── edit.html                # Edição de registro
│   └── placeholders.html        # Template dos emails
├── Registro/                    # Banco de dados local
│   └── Registro.xlsx            # Planilha com histórico das malas
└── uploads/                     # Pastas criadas para cada mala
    └── [nome_arquivo] [folha]/  # Arquivos salvos por campanha

Fluxo de Uso
1. Na página inicial, clique em "Enviar Nova Mala Direta"
2. Selecione o arquivo Excel ou CSV
3. Escolha a planilha desejada
4. Marque as colunas que serão incluídas no email
5. Defina qual coluna contém o email e qual contém o nome do destinatário
6. Escolha a origem dos dados (arquivo, dataframe, ambos ou nenhum)
7. Digite o assunto e componha o corpo do email
8. Se desejar, marque o modo de teste para validar antes do envio real
9. Clique em "Enviar Mensagens"

Para reutilizar uma mala existente, acesse o histórico, duplique o registro desejado e faça as adaptações necessárias.

Segurança
- Apenas o criador do registro pode editá-lo
- Todos os membros da equipe podem visualizar e enviar qualquer mala
- A duplicação cria um novo registro
- Verificação de duplicatas evita registros repetidos

Desenvolvido por Alysson Uchoa
