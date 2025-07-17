# web_rel

## SUMARY

1. Introdução
    - Descrição do projeto
    - Objetivo do projeto
2. Requisitos
    - Funcionais
    - Não funcionais
    - Dependências
3. Explicação do código
    - Funções
4. Contato

## Introdução

<p style="text-align: justify;">O código do qual essa documentação trata tem por objetivo documentar para usos futuros o código web_rel, mas também tem por objetivo servir como um possível referencial para desenvolvimentos semelhantes em automação em navegadores. Por isso, quando tratar dos usos, poderá ser observado uma explicação mais detalhada sobre o que cada elemento faz.<br>
Uma parte dos elementos não é representada dentro do código a seguir para manter o sigilo e o resguardo dos dados da instituição para a qual foi desenvolvida a aplicação.<p>

### Objetvo do projeto

<p style="text-align: justify;">O código foi criado baseado num problema que foi encontrado ao iniciar a atividade de acompanhamento de geração de ordens de serviço, seria necessário baixar todos os dias, para alimentar uma planilha de acompanhamento que fica disponível para toda equipe, um relatório de anotações realizadas. Como ter que acessar o site, navegar, baixar o relatório, extrair os dados e editar a planilha final, além de remover a formatação CSV, esse processo diário seria trabalhoso. Por isso esse código foi desenvolvido, para automatizar esta tarefa de forma completa, um dos complementos não citados no código é que ele foi configurado para executar nas tarefas diárias do computador que o executará.<p>

### Descrição do projeto
<p style="text-align: justify;">O código é desenvolvido em linguagem Python com forte uso de selenium para automatizar os processos de navegação. Em grande parte o uso dos caminhos XPATH foram utilizados para que fosse mais objetivo o acesso a cada elemento, já que o site não apresentava padronização em alguns elementos, sendo o XPATH a melhor solução para criar uma função mais direta e finalizar o código de forma limpa.<br>
Havia uma versão anterior, criada de forma improvisada, esta possuia diversos elementos para navegação, mas a legibilidade era comprometida e a complexidade era maior. Por isso, a melhor forma foi usar o XPATH para ordenar isso. Ainda sobre a versão primária deste código, ela era muito expositiva quanto a dados sensíveis para subir ao github, por isso os elementos nesse novo modelo são armazenados dentro de um arquivo JSON.<br>
Sobre a versão atual, além da navegação em navegador web, também realiza a operação de edição de uma planilha, inserindo valores a partir de um relatório em formato CSV, para manter a base que a equipe usa para trabalhar atualizada. Por fim, a partir desse código é gerado um executável para que o responsável possa exeutar em sua máquina, tendo também um arquivo para configurar a ativação das atividades diárias e permitir a execução desta aplicação sem a necessidade dele lembrar todo dia.<p>

## Requisitos

### Requisitos funcionais

- <p style="text-align: justify;">A aplicação deve acessar o navegador;<p>
- <p style="text-align: justify;">A aplicação deve realizar o login no sistema;<p>
- <p style="text-align: justify;">A aplicação deve navegar no site até o local onde estão os relatórios;<p>
- <p style="text-align: justify;">A aplicação deve baixar o relatório pré-estabelecido;<p>
- <p style="text-align: justify;">A aplicação deve acessar o relatório baixado;<p>
- <p style="text-align: justify;">A aplicação deve ler o relatório;<p>
- <p style="text-align: justify;">A aplicação deve abrir a planilha excel;<p>
- <p style="text-align: justify;">A aplicação deve excluir os dados anteriores; e<p>
- <p style="text-align: justify;">A aplicação deve inserir os dados extraídos do relatório.<p>

### Requisitos não-funcionais

- <p style="text-align: justify;">A aplicação deve rodar em máquinas com MAC OS ou Windows OS;<p>
- <p style="text-align: justify;">A aplicação deve rodar no Chrome;<p>
- <p style="text-align: justify;">A aplicação deve ser capaz de acessar arquivo JSON e XLSX;<p>
- <p style="text-align: justify;">A aplicação deve ser executada sem comprometer a atividade de trabalho do usuário, ou seja, não pode abrir uma janela do navegador ao inicialiar.<p>

### Dependências

1. Bibliotecas:
    - <p style="text-align: justify;">Selenium: utilizado para navegar e automatizar operações no navegador<p>
    - <p style="text-align: justify;">Pathlib: utilizando o Path para manipulação de caminhos sem depender de formatação de acordo com o sistema operacional (OS)<p>
    - <p style="text-align: justify;">Json e CSV: utilizados para acessar os respectivos tipos de documentos<p>
    - <p style="text-align: justify;">Time: utilizando o sleep para determinar um tempo de pausa para execução dos comandos de navegação<p>
    - <p style="text-align: justify;">Datetime: utilizando o date para coletar a data do dia e permitir a localização do arquivo de forma dinâmica<p>
    - <p style="text-align: justify;">Openpyxl: utilizado para manipular o arquivo XLSX.<p>

## Explicação do código

### Funções

1. Leitura de arquivos JSON:
    - <p style="text-align: justify;">Abre arquivos JSON no modo leitura e retorna seus dados.<p>

    ```python
    def read_json(path_arg=""):
    with open(path_arg, 'r') as file:
        json_data = json.load(file)
        return json_data 
    ```
2. Criando instância do navegador
    - <p style="text-align: justify;">Essa função cria uma instância do navegador do Chrome, retornando um objeto do tipo webdriver.chrome, configurando-o. No caso do código abaixo não foi relizado, mas no código enviado ao responsável, foi habilitada a opção de abrir o navegador sem gerar uma janela.<br>
    Como explicado de forma superficial e implicita acima, a variável option carrega as opções do navegador, suas configurações de uso.<br>
    A variável services carrega o serviço do navegador, ou seja, a partir de qual ferramenta ele será aberto, por default o selenium pode abrir um navegador e realizar operações sem o driver, mas é fundamental ter um.<br><br>
    IMPORTANTE: EM CASO DE ERRO, ATUALIZAR O CHROME DRIVER.<br><br>
    Por fim, o retorno dessa função, como explicado acima é um objeto webdriver.chrome já configurado com as opções e serviço idealizados para o leno funcionamento do código.
    <p>

    ```python
    def make_browser() -> webdriver.Chrome:
    options = webdriver.ChromeOptions()
    services = Service(executable_path=CHROME_DRIVER_EXE)
    browser = webdriver.Chrome(options=options, service=services)
    return browser
    ```
3. Definindo um elemento clicável
    - <p style="text-align: justify;"><p>
    ```python
    def make_click(arg_local=""):
    local_click = nav_elements[arg_local]
    element_clickable = EC.element_to_be_clickable
    clickable = WebDriverWait(browser, 12.0).until(
        element_clickable(
            (By.XPATH, local_click)
        )
    )
    return clickable
    ```