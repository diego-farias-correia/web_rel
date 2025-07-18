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
    - Navegação web
        - Constantes e arquivos JSON
            - Constantes
            - Arquivos JSON
        - Funções
            - Carregando arquivos JSON
            - Criando instância do navegador
            - Definindo elemento clicável
        - Estruturas de repetição e uso do select
            - Estruturas de repetição for
                - for: login
                - for: navegação
            - Select
    - Manipulação de relatório
        - Arquivo CSV
        - Arquivo XLSX
            
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

### Navegação web

#### Constantes e arquivos JSON

1. Constantes
    - <p style="text-align: justify;">As constantes abaixo servem como fundamentos para acesso das principais operações desse código.<br>
    - Root_FOLDER recebe o caminho da pasta principal onde o código roda. IMPORTANTE: NO EXECUTÁVEL, O __file__ NÃO É ÚTIL, SENDO NECESSÁRIO IMPORTAR O MÓDULO SYS PARA PASSAR PARA O PATH __"sys.executable"__, que retornará o caminho para o executável.<br>
    - CONFIG e NAV_KEY são arquivos JSON que serão melhor abordados abaixo no tópico "arquivos JSON"<p>
    ```python
    ROOT_FOLDER = Path(__file__).parent
    CHROME_DRIVER_EXE = ROOT_FOLDER / "chromedriver-win64" / "chromedriver.exe"
    CONFIG = ROOT_FOLDER / "config.json"
    NAV_KEY = ROOT_FOLDER / "nav_key.json"
    ```
2. Arquivos JSON
    - <p style="text-align: justify;">Os arquivos JSON chamados config e nav_key, são arquivos que carregam em si informações para login e navegação respectvamente.<br>
    - Config recebe os dados do input do usuário, nos quais, um é o recebimento do usuário e senha, gurdados em formato dict e o outro é o caminho para a pasta do arquivo XLSX. Essas informações são carregadas para o JSON com o json.dump, quando não existe o arquivo config e existindo, são carregadas através do json.load.<br>
    - Nav_key contém um dicionário com todos os XPATH necessários para os elementos, desde o login, até o final no botão de download do relatório.<p>
    Caso não exista o arquivo:
    ```python
    if not CONFIG.exists():
    input_user = input("Insert your user name (name.lastname): ")
    input_password = input("Insert your password: ")
    print("WARNING: don't insert the file, only the path to folder!")
    xlsx_local = input("Insert the path to file xlsx: ")

    dict_dump = {
        "nome-usuario": input_user,
        "password": input_password
    }

    with open(CONFIG, 'w') as file:
        json.dump((dict_dump, xlsx_local), file)
    ```
    Caso o arquivo já exista, o código a ser executado encontra-se abaixo na seção funções.

#### Funções

1. Carregando arquivos JSON:
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
    - <p style="text-align: justify;">Essa função utiliza o módulo expected_conditions (EC) e a classe WebDriverWait para esperar por elementos e retornar um objeto HTML clicável.<br>
    A variável local_click recebe do dicionário nav_elements, o XPATH condizente a chave passada pelo argumento arg_local.<br>
    A variável element_clickable recebe a função do módulo EC, element_to_be_clickable, que é uma função que designa o que deve ser esperado.<br>
    A variavel clickable recebe um objeto WebElement, que nesse caso é m oca clicável que permite a operação click(). No entanto, para  isso, ele usa a classe WebDriverWait para esperar que o elemento localizado por, agora carregado nele, element_clickable e assim retornar o objeto.<p>
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

#### Estruturas de repetição e uso do select

1. Laços de repetição for
    1. Laço for: login
        - <p style="text-align: justify;">A estrutura de repetição abaixo visa consumir o dicionário de login, para inserir o nome de usuário e a senha. Como são realizados mais processos nessaa etapa, como ddigitar e pressionar Enter, essa estrutura foi separada da estrutura for a seguir.<p>
        ```python
        for search_id, data_user in login_elements.items():
            login = make_click(search_id)
            login.click()
            login.send_keys(data_user)
            login.send_keys(Keys.ENTER)
        ```
    2. Laço for: navegação
        - <p style="text-align: justify;">A estrutura de repetição abaixo visa consumir todo o dicionário de navegação, este dicionário possui todos os XPATH, inclusive com elementos que já foram trabalhados previamente no código, por isso a necessidade de criar uma estrutura condicional para pulr estes laços - embora eu possa reconhecer que uma forma melhor talvez fosse remover essas chaves da lista em um momento anterior no código.<br>
            A estrutura segue de forma linear até selecionar e baixar o arquivo do relatório desejado.<p>
        ```python
        for click_operation in nav_elements:
            if click_operation in ["site", "nome-usuario", "password"]:
                continue
            navegation = make_click(click_operation)
            navegation.click()
        ```
2. Select
    - <p style="text-align: justify;">O uso do select, por ser casual não foi montado dentro da mesma função, já conhecida, make_click, mas foi adaptada aqui por ter um processo distinto. Ele utiliza a classe select que recebe um objeto WebElement para retornar as opções da caixa de seleção, por sua vez, ele buscará o elemento através do find_element, que é um método do webdriver.chrome, tornando possível a seleção do texto visível esperado.<p>
    ```python
    perfil = Select(
        WebDriverWait(browser, 12.5).until(
            EC.presence_of_element_located(
                (By.TAG_NAME, "select")
            )
        )
    )
    selection = browser.find_element(By.XPATH, "//button[contains(text(), 'Selecionar')]")
    perfil.select_by_visible_text("PRD-SBT-Gestor")
    selection.send_keys(Keys.ENTER)
    ```

### Manipulação de relatório

#### Arquivo CSV

<p style="text-align: justify;">A manipulação do arquivo csv é resumida a extração dos dados do arquivo do relatório baixado, isso é feito com o uso da biblioteca csv para que com o uso do csv reader que retorna um iterator, pelo qual é possível criar uma lista na forma em que está posta para iterar sobre a aba no documento XLSX.<p>

```python
with open(csv_file, 'r', encoding="utf-8") as file:
    data = list(csv.reader(file))
```

#### Arquivo XLSX

<p style="text-align: justify;">Esse trecho do código faz a criaçao de um objeto workbook que carrega em si todo um arquivo excel, dentro desse arquivo a variável table captura a aba L.A, onde as linhas são limpas para que novos valores possam ser inseridos na estrutura condicional lgo abaixo.<p>

```python
workbook = load_workbook(str(xlsx_file))
table = workbook["L.A"]
table.delete_rows(1, table.max_row)

for index_line, line in enumerate(data, start=1):
    for index_column, data_value in enumerate(index_line, start=1):
        table.cell(row=index_line, column=index_column, value=data_value)
```

## Contato