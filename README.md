# 🛒 Extração dos Mais Vendidos — Amazon Brasil (Computadores e Informática)

Este projeto é uma automação web feita em Python que navega pelos menus da Amazon Brasil, entra no departamento de **Computadores e Informática** e coleta os itens do carrossel de **Mais Vendidos**. 

O objetivo principal aqui foi construir um fluxo que não apenas coletasse os dados (título, preço, nota, total de avaliações e o link do produto), mas que também tratasse esses dados e gerasse uma planilha Excel limpa, bonita e pronta para análise, sem que o usuário precise formatar células na mão depois.

Além da extração em si, adicionei um monitor de performance para acompanhar o tempo de execução e o consumo de memória RAM durante o processo.

---

## 💡 Desafios Técnicos e Soluções Desenvolvidas

Quem já tentou automatizar a Amazon sabe que o site é cheio de pegadinhas para robôs. Durante o desenvolvimento, encontrei três desafios principais e resolvi da seguinte forma:

* **O Menu Deslizante e o Clique Interceptado:** Quando clicamos no menu "Todos" (o hambúrguer lá no canto superior esquerdo), a Amazon abre uma gaveta com animação CSS. Se tentarmos dar um clique normal do Selenium no item de "Computadores e Informática", quase sempre tomamos um erro de `ElementClickInterceptedException` porque o painel ainda está se mexendo na tela.
  * **Solução:** Em vez de depender do clique físico do mouse (que faz teste de coordenadas na tela), forcei o clique direto via JavaScript (`arguments[0].click()`). Assim, o evento vai direto no nó do DOM e ignora qualquer animação ou tempo de carregamento da interface.
* **Itens Fora da Tela (Offscreen) e Lazy Loading:** No carrossel de Mais Vendidos, apenas os primeiros itens aparecem visíveis. Se você pede para o Selenium ler o texto de um produto que está escondido à direita, ele simplesmente retorna uma string vazia `""` porque o elemento não está visível.
  * **Solução:** Dentro do loop de extração, adicionei um comando para deslizar o carrossel na horizontal (`scrollIntoView({inline: 'center'})`), centralizando item por item na tela. Além disso, troquei a leitura padrão do Selenium pelo atributo `textContent`, que consegue pegar o texto no DOM mesmo se a Amazon esconder o preço em tags para leitores de tela (como a classe `a-offscreen`).
* **Dados Limpos no Excel (Nada de texto puro):** É muito comum extrair dados da internet e jogar tudo como texto no Excel. O problema é que "R$ 1.056,00" ou "(1.420)" como texto impedem qualquer cálculo de média ou soma.
  * **Solução:** Criei funções auxiliares que limpam as strings e convertem tudo para números reais (`float` e `int`). Na hora de salvar com o `openpyxl`, aplico a formatação nativa de moeda e casas decimais nas próprias células, coloco um link clicável bem discreto ("Ver na Amazon") e ainda adiciono uma linha de resumo no final usando fórmulas nativas do Excel (`=AVERAGE` e `=SUM`).

---

## 🛠️ Tecnologias Utilizadas

Para fazer tudo isso funcionar sem dor de cabeça, utilizei as seguintes bibliotecas:

* **selenium:** Para controlar o navegador, clicar nos menus e deslizar o carrossel.
* **webdriver-manager:** Para baixar e gerenciar o ChromeDriver automaticamente. Você não precisa se preocupar em baixar o driver na mão nem se a versão do seu Chrome atualizou.
* **openpyxl:** Para criar e estilizar a planilha do Excel (cores, bordas, fontes, congelamento de painel e fórmulas).
* **tracemalloc e logging (nativos do Python):** Utilizados no nosso decorador de performance para medir exatamente quanto tempo o script levou e qual foi o pico de memória RAM utilizada.

---

## 📁 Estrutura de Pastas

Organizei o código de uma forma bem modular para separar a lógica de extração das ferramentas de medição:

```text
├── include/
│   ├── functions.py       # Funções de extração (Selenium), limpeza de dados e exportação (Excel)
│   └── performance.py     # Nosso decorador @perf_tracker para medir tempo e memória
├── main.py                # Arquivo principal que configura o navegador e roda o fluxo
├── requirements.txt       # Lista de dependências do projeto
├── README.md              # Documentação do projeto
└── Top_Mais_Vendidos_Amazon.xlsx  # Planilha gerada após a execução
```

---

## ⚙️ Configuração do Ambiente

Aqui está o passo a passo padrão para você rodar o projeto isolado em um ambiente virtual no seu computador:

**1° Criar e ativar o Ambiente Virtual:**
Isso é importante para não misturar as bibliotecas desse script com outros projetos que você tenha no computador.

```bash
# Criar o ambiente
python -m venv .venv

# Ativar no Linux ou macOS
source .venv/bin/activate

# Ativar no Windows (pelo terminal padrão / CMD ou PowerShell)
.\.venv\Scripts\activate 
```

**2° Instalar as Dependências:**
Com o ambiente ativo, instale os pacotes listados no arquivo `requirements.txt`:

```bash
pip install -r requirements.txt
```

*(Se você ainda não criou o arquivo `requirements.txt`, ele só precisa ter o seguinte conteúdo: `selenium`, `webdriver-manager` e `openpyxl`).*

---

## 🚀 Como Executar o Código

Com tudo instalado, basta rodar o arquivo principal pelo terminal:

```bash
python main.py
```
*(ou `python3 main.py` dependendo de como está configurado no seu sistema).*

### Quero extrair mais (ou menos) itens, como faço?
Por padrão, o script está configurado para pegar os **10 primeiros itens** do carrossel. Se você quiser mudar esse limite, é só abrir o `main.py` e alterar o parâmetro `max_items` na chamada da função:

```python
# Exemplo para pegar os Top 15 produtos
results = extract_best_sellers(driver, target_url, max_items=15)
```

---

## ⏱️ Performance e Benchmarking

Para garantir que o código não está gastando processamento ou memória à toa, criei o decorador `@perf_tracker` que fica no arquivo `include/performance.py`. Sempre que a função principal termina de rodar, ela imprime um log no terminal mostrando os custos computacionais daquela extração.

No final da execução, você verá um retorno no terminal mais ou menos assim:

```text
INFO:__main__:--- extract_best_sellers ---
INFO:__main__:Execution time: 4.8210 seconds
INFO:__main__:Peak memory usage: 1.2450 MB
INFO:__main__:Excel file successfully generated and saved as: Top_Mais_Vendidos_Amazon.xlsx
```

Como o processamento do DOM é rápido e delegamos o limite de itens direto para o motor do XPath (`position() <= X`), o pico de consumo de memória (RAM) costuma ficar super leve (em torno de apenas 1 a 2 MB), mesmo abrindo uma instância real do navegador!

---

## 📊 O Que Você Terá no Final?

Assim que o script fechar o navegador, você encontrará o arquivo `Top_Mais_Vendidos_Amazon.xlsx` na pasta raiz do projeto. A planilha final já vem formatada com:
* Título principal com a data e hora exatas da extração.
* Linhas zebradas (cinza claro e branco) para facilitar a leitura.
* Preços formatados em Real (`R$`) e avaliações formatadas em estrelas (`★`).
* Uma coluna com o link **"Ver na Amazon"** clicável, levando direto para a página do produto.
* Uma linha no final com a **média de preço**, a **média de estrelas** e a **soma de todas as avaliações** dos produtos que você extraiu.